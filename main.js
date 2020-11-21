const fs = require("fs");
var exec = require('child_process').exec;
var parser = require('fast-xml-parser');
const uuid = require("uuid");
const path = require("path");

let typePrefix = "SB";

function execute(command){
    return new Promise(((resolve, reject) => {
        exec(command, function(error, stdout, stderr){
            if(error) reject(error);
            else resolve(stdout);
        });
    }))
};

async function getXML(zipFile, subFile) {
    const xml = await execute("unzip -qc '"+ zipFile + "' " + subFile);
    return parser.parse(xml,{
        ignoreAttributes : true,
        trimValues: false,
    });
}

async function extract(file) {
    const matches = path.basename(file).match(/(sb|ph)([0-9]+)/i);
    const name = matches[2].padStart(3, "0");

    const prez = await getXML(file, "ppt/presentation.xml");
    const nSlides = prez["p:presentation"]["p:sldIdLst"]["p:sldId"].length;

    let title = "";
    let author = "";
    let year = "";
    let verses = [];

    for(let i = 0; i < nSlides; i++) {
        const obj = await getXML(file, "ppt/slides/slide" + (i+1) + ".xml");

        const contentRoot = obj["p:sld"]["p:cSld"]["p:spTree"]["p:sp"];

        const content = contentRoot.map(row => {
            if ("p:txBody" in row) {
                return textR(row).replace(/\n{2,}/g, "\n").trim();
            }

            return null;
        }).filter(r => !!r);

        if(content.length === 0) continue;

        if(i === 0) {
            title = content[1];
            const meta = extractMeta(content[2]);
            author = meta.author;
            year = meta.year;
        }

        verses.push({
            content: sanitizeContent(content[0]),
            verseType: verseType(content),
        });
    }

    return {
        name,
        title,
        author,
        year,
        verses,
    }
}

function verseType(content) {
    const options = content.filter(c =>  {
        if(c.length >= 4 && c.length < 50 && c.match(/[0-9]+/)) return true;
        return c.startsWith("SB");
    });
    const minLen = Math.min(...options.map(l => l.length));
    const verseId = options.filter(c => c.length === minLen)[0];

    if(verseId === undefined) {
        if(content.filter(r => r === "Refrain").length === 1) return "refrain";
    }

    if(!verseId) console.log(content);

    if(verseId.match(/v[0-9 ab]+$/i)) {
        return "verse";
    }

    if(verseId === "SB64 V1-2") return "verse";

    if(verseId.match(/refrain/i)) return "chorus";
    if(verseId.match(/chorus/i)) return "chorus";
    if(verseId.match(/repeat/i)) return "chorus";
    if(verseId.match(/bridge/i)) return "bridge";
    if(verseId.match(/ending/i)) return "ending";
    if(verseId.match(/End/i)) return "ending";

    console.log("no-verse-id-match", content);
    return "chorus";
}

function sanitizeContent(content) {
    return content.replace(/~/g, "").trim().replace(/\n.$/, ".").replace(/\n,\n/g, ",\n").replace(/\n{2,}/g, "\n");
}

function extractMeta(value) {
    let year = "";
    let author = "";

    switch(value) {
        case "“From the Depths My Prayer Ascendeth Ethelbert W. Bullinger, 1877":
            return {author: "Ethelbert W. Bullinger", year: "1877"};
        case "“By Babel’s Streams We Sat and WeptWilliam B. Bradbury, 1853":
            return {author: "William B. Bradbury", year: "1853"};
        case "“With All My Heart Will I RecordLouis Bourgeois, 1543":
            return {author: "Louis Bourgeois", year: "1543"};
        case "“To God My Earnest Voice I RaiseLowell Mason, 1824":
            return {author: "Lowell Mason", year: "1824"};
        case "“Father, Again in Jesus’ Name We MeetJames Langran, 1862":
            return {author: "James Langran", year: "1862"};
        case "\"The Ends of all the Earth Shall Hear“Composed by: William H. Doane":
            return {author: "William H. Doane", year: ""};
        case "“I Sought the Lord, and Afterward I KnewJean Silbelius, 1865-1957":
            return {author: "Jean Silbelius", year: "1957"};
    }

    if(value.indexOf("\n") !== -1) {
        value = value.split("\n")[1];
    } else if(value.indexOf('”') !== -1) {
        value = value.substr(value.indexOf('”')+1);
    } else if(value.lastIndexOf('"') !== -1) {
        value = value.substr(value.lastIndexOf('"')+1);
    }

    const yearMatch = value.match(/[0-9]{4}$/);
    if(yearMatch) year = yearMatch[0];
    author = value
        .replace(/[0-9]{4}.*$/, "")
        .replace(" (v5-7)Music arrgd by:", ",")
        .replace(/&amp;/g, "&")
        .replace(/(^| )(words|by|music|translated|and|from|arranged|lyrics|arrng|arrg|adapt)/ig, "")
        .replace(/(music|words)/ig, "")
        .replace(/[;:]/g, ',')
        .replace(/[ ]{2,}/g, ' ')
        .replace(/(\.,|,\.)/g, ',')
        .replace(/[,]{2,}/g, ',')
        .replace(/[ ,]+$/, "")
        .replace(/^[ ,]+/, "")
        .replace(/ ,/g, ",")
        .trim();

    if(author.indexOf('PH') !== -1 || author.indexOf("SB") !== -1)
        return {author: "", year: ""};

    return {
        year,
        author,
    }
}

function sanitizeRtfContent(value) {
    return value.replace(/’/g, "\\'92").replace(/\n/g, '\\\n');
}

function toBase64(value) {
    return Buffer.from(value).toString("base64");
}

function textR(value, debug = false, indent = 0) {
    const indentStr = " ".repeat(indent);
    if (typeof value === "string") {
        if(debug) console.log(indentStr + value, value.replace(/[\x00-\x1F\x7F-\x9F]/g, "<CTRL>"));
        if(value === "$") return;
        return value;
    }

    if (value instanceof Array) {
        return value.map(v => textR(v, debug, indent + 1)).join("");
    }

    let list = [];

    for (let i in value) {
        if(debug) console.log(indentStr + i);
        if(i === "lang") continue;
        if(i === "dirty") continue;
        if(i === "a:endParaRPr") {
            list.push("\n");
            continue;
        }

        list.push(textR(value[i], debug, indent + 1));

        if(i === "a:r") list.push("\n");
    }

    return list.join("");
}

function blankSlide() {
    return `<RVSlideGrouping color="0 0 0 0" name="" uuid="${newUUID()}">
            <array rvXMLIvarName="slides">
                <RVDisplaySlide UUID="${newUUID()}" backgroundColor="0 0 0 1" chordChartPath="" drawingBackgroundColor="false" enabled="true" highlightColor="0.9859483242034912 0 0.02695056796073914 1" hotKey="" label="Blank Slide" notes="" socialItemCount="1">
                    <array rvXMLIvarName="cues"></array>
                    <array rvXMLIvarName="displayElements"></array>
                </RVDisplaySlide>
            </array>
        </RVSlideGrouping>`;
}

function grouping(name, children) {
    return `<RVSlideGrouping color="0 0 0.9981889724731445 1" name="${name}" uuid="${newUUID()}">
            <array rvXMLIvarName="slides">${children}</array>
        </RVSlideGrouping>`
}

function verseContentText(verse) {
    // console.log(verse + "\n\n");
    verse = sanitizeRtfContent(verse);

    return toBase64(`{\\rtf1\\ansi\\ansicpg1252\\cocoartf2513
\\cocoatextscaling0\\cocoaplatform0{\\fonttbl\\f0\\fswiss\\fcharset0 Helvetica;}
{\\colortbl;\\red255\\green255\\blue255;}
{\\*\\expandedcolortbl;;}
\\deftab720
\\pard\\pardeftab720\\qc\\partightenfactor0

\\f0\\fs180 \\cf1${verse}}\\cf0`);
}

function titleText(title) {
    title = sanitizeRtfContent(title);
    const text = `{\\rtf1\\ansi\\ansicpg1252\\cocoartf2513
\\cocoatextscaling0\\cocoaplatform0{\\fonttbl\\f0\\fswiss\\fcharset0 Helvetica-Bold;}
{\\colortbl;\\red255\\green255\\blue255;}
{\\*\\expandedcolortbl;;}
\\deftab720
\\pard\\pardeftab720\\qc\\partightenfactor0

\\f0\\b\\fs180 \\cf1  ${title}}`;

    return toBase64(text);
}

function verseText(verse) {
    verse = sanitizeRtfContent(verse);
    const text = `{\\rtf1\\ansi\\ansicpg1252\\cocoartf2513
\\cocoatextscaling0\\cocoaplatform0{\\fonttbl\\f0\\fswiss\\fcharset0 Helvetica;}
{\\colortbl;\\red255\\green255\\blue255;}
{\\*\\expandedcolortbl;;}
\\deftab720
\\pard\\pardeftab720\\qr\\partightenfactor0

\\f0\\fs120 \\cf1 ${verse}}`;
    return toBase64(text);
}

function slide(title, verse, content) {
    return `<RVDisplaySlide UUID="${newUUID()}" backgroundColor="0 0 0 1" chordChartPath="" drawingBackgroundColor="false" enabled="true" highlightColor="0 0 0 0" hotKey="" label="" notes="" socialItemCount="1">
                    <array rvXMLIvarName="cues"></array>
                    <array rvXMLIvarName="displayElements">
                        <RVTextElement UUID="${newUUID()}" additionalLineFillHeight="0.000000" adjustsHeightToFit="false" bezelRadius="0.000000" displayDelay="0.000000" displayName="TextElement" drawLineBackground="false" drawingFill="false" drawingShadow="false" drawingStroke="false" fillColor="" fromTemplate="false" lineBackgroundType="0" lineFillVerticalOffset="0.000000" locked="false" opacity="1.000000" persistent="false" revealType="0" rotation="0.000000" source="" textSourceRemoveLineReturnsOption="false" typeID="0" useAllCaps="false" verticalAlignment="0">
                            <RVRect3D rvXMLIvarName="position">{991 1042 0 852 80}</RVRect3D>
                            <shadow rvXMLIvarName="shadow">0.000000|0 0 0 1|{4.94974746830583, -4.94974746830583}</shadow>
                            <dictionary rvXMLIvarName="stroke">
                                <NSColor rvXMLDictionaryKey="RVShapeElementStrokeColorKey">0 0 0 1</NSColor>
                                <NSNumber hint="float" rvXMLDictionaryKey="RVShapeElementStrokeWidthKey">1.000000</NSNumber>
                            </dictionary>
                            <NSString rvXMLIvarName="RTFData">${verseText(verse)}</NSString>
                        </RVTextElement>
                        <RVTextElement UUID="${newUUID()}" additionalLineFillHeight="0.000000" adjustsHeightToFit="false" bezelRadius="0.000000" displayDelay="0.000000" displayName="TextElement" drawLineBackground="false" drawingFill="false" drawingShadow="false" drawingStroke="false" fillColor="1 1 1 1" fromTemplate="false" lineBackgroundType="0" lineFillVerticalOffset="0.000000" locked="false" opacity="1.000000" persistent="false" revealType="0" rotation="0.000000" source="" textSourceRemoveLineReturnsOption="false" typeID="0" useAllCaps="false" verticalAlignment="1">
                            <RVRect3D rvXMLIvarName="position">{67 58 0 1813 230}</RVRect3D>
                            <shadow rvXMLIvarName="shadow">0.000000|0 0 0 0.3294117748737335|{4, -4}</shadow>
                            <dictionary rvXMLIvarName="stroke">
                                <NSColor rvXMLDictionaryKey="RVShapeElementStrokeColorKey">0 0 0 1</NSColor>
                                <NSNumber hint="float" rvXMLDictionaryKey="RVShapeElementStrokeWidthKey">1.000000</NSNumber>
                            </dictionary>
                            <NSString rvXMLIvarName="RTFData">${titleText(title)}</NSString>
                        </RVTextElement>
                        <RVTextElement UUID="${newUUID()}" additionalLineFillHeight="0.000000" adjustsHeightToFit="false" bezelRadius="0.000000" displayDelay="0.000000" displayName="Default" drawLineBackground="false" drawingFill="false" drawingShadow="false" drawingStroke="false" fillColor="1 1 1 1" fromTemplate="false" lineBackgroundType="0" lineFillVerticalOffset="0.000000" locked="false" opacity="1.000000" persistent="false" revealType="0" rotation="0.000000" source="" textSourceRemoveLineReturnsOption="false" typeID="0" useAllCaps="false" verticalAlignment="0">
                            <RVRect3D rvXMLIvarName="position">{82 2 0 1755 1195}</RVRect3D>
                            <shadow rvXMLIvarName="shadow">0.000000|0 0 0 1|{4, -4}</shadow>
                            <dictionary rvXMLIvarName="stroke">
                                <NSColor rvXMLDictionaryKey="RVShapeElementStrokeColorKey">0 0 0 1</NSColor>
                                <NSNumber hint="double" rvXMLDictionaryKey="RVShapeElementStrokeWidthKey">0.000000</NSNumber>
                            </dictionary>
                            <NSString rvXMLIvarName="RTFData">${verseContentText(content)}</NSString>
                        </RVTextElement>
                    </array>
                </RVDisplaySlide>`
}

function proPresenterXmlTemplate(options) {

    const {author, year, title, verses} = options;
    let verseCounter = 0;

    console.log(options.name);

    const groups = verses.map((v, index) => {
        const slideTitle = index === 0 ? title : "";

        let groupName = "";

        switch(v.verseType) {
            case "chorus":
                groupName = "Chorus";
                break;
            case "bridge":
                groupName = "Bridge";
                break;
            case "ending":
                groupName = "Ending";
                break;
            case "verse":
                verseCounter++;
                groupName = "Verse " + verseCounter;
                break;
        }

        const verseName = typePrefix + " " + options.name.replace(/^[0]{1,}/g, "") + " " + groupName.toLowerCase();

        const lines = v.content.split("\n");
        let parts = [];
        if(lines.length > 7 || v.content.length > 260) { // split big verses
            console.log(groupName, (index + 1), "is split over 2 slides");
            const n = Math.ceil(lines.length / 2);
            const part0 = lines.slice(0, n);
            const part1 = lines.slice(n);

            parts.push(slide(slideTitle, verseName, part0.join("\n")))
            parts.push(slide("", verseName, part1.join("\n")))
        } else {
            parts.push(slide(slideTitle, verseName, lines.join("\n")));
        }

        return grouping(groupName, parts.join("\n"));
    })

    return `<RVPresentationDocument CCLIArtistCredits="${author}" CCLIAuthor="" CCLICopyrightYear="${year}" CCLIDisplay="true" CCLIPublisher="" CCLISongNumber="" CCLISongTitle="${title}" backgroundColor="0 0 0 0" buildNumber="100991490" category="Presentation" chordChartPath="" docType="0" drawingBackgroundColor="false" height="1200" lastDateUsed="2020-11-03T20:44:45-05:00" notes="" os="2" resourcesDirectory="" selectedArrangementID="" usedCount="0" uuid="${newUUID()}" versionNumber="600" width="1920">
    <RVTimeline duration="0.000000" loop="false" playBackRate="1.000000" rvXMLIvarName="timeline" selectedMediaTrackIndex="0" timeOffset="0.000000">
        <array rvXMLIvarName="timeCues"></array>
        <array rvXMLIvarName="mediaTracks"></array>
    </RVTimeline>
    <array rvXMLIvarName="groups">
        ${blankSlide()}
        ${groups}
    </array>
    <array rvXMLIvarName="arrangements"></array>
</RVPresentationDocument>`
}

function newUUID() {
    return uuid.v4().toUpperCase();
}

async function pptx2pro(file) {
    const options = await extract(file);

    if(options.verses.length === 0) {
        console.log(typePrefix + options.name, "is blank");
    }

    const xml = proPresenterXmlTemplate(options);
    const outputName = options.name + ".pro6";

    fs.writeFileSync("./output/"+outputName, xml);
    // console.log("\tcompleted " + outputName);
}

async function process(prefix) {
    typePrefix = prefix.toUpperCase();

    console.log("setting up output directory...");

    if(fs.existsSync(`./bundle-${prefix}.pro6x`)) {
        fs.unlinkSync(`./bundle-${prefix}.pro6x`);
    }

    if(fs.existsSync("./output")) {
        fs.rmdirSync("./output", {
            recursive: true,
        });
    }

    fs.mkdirSync("./output");

    console.log("discovering src files");
    const srcDir = "./src-" + prefix;
    let files = fs.readdirSync(srcDir)

    console.log("processing files...");

    var todo = [];

    for(let i = 0; i < files.length; i++) {
        // console.log(i, " of ", files.length);
        todo.push(pptx2pro(srcDir + "/" + files[i]));
    }

    await Promise.all(todo);

    console.log("creating bundle...");
    await exec(`zip bundle-${prefix}.pro6x ./output/*.pro6`);
    console.log("done!");
}

async function main() {
    await process("ph");
    await process("sb");
}

main();