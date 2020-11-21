
# PPTX Extractor

Converts our church's specific format of pptx songs into pro-presenter-6 bundles.

## Usage

Requires Node and source pptx files (see sample-source-file for expected format)

Expects source directories  
```
./src-ph
    # Psalter hymnal pptx files with names like PH-###-{name}.pptx
./src-sb
    # Song book pptx files with names like SB-###-{name}.pptx
```

Run program
```
> npm install
> node main.js
    # outputs bundle-ph.pro6x and bundle-sb.pro6x
```

## FAQ

### Warning!! PPTX decoding is not perfect
I wasn't able to find a reliable pptx decoding library that dealt with things like line breaks and
special characters properly. This is a best attempt to decode things properly.

### Why pro6x format?

The newer pro-presenter format uses protobuf which is much harder to reverse-engineer. 
This was easier.