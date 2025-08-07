# officeParser-min
A Node.js library to parse text out of any office file. 
A fork from harashhankur/officeParse but without bulky PDF parsing features.

### Supported File Types

- [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`odt`](https://en.wikipedia.org/wiki/OpenDocument)
- [`odp`](https://en.wikipedia.org/wiki/OpenDocument)
- [`ods`](https://en.wikipedia.org/wiki/OpenDocument)



## Install via npm

```
npm i officeparser
```

## Command Line usage
If you want to call the installed officeParser.js file, use below command
```
node <path/to/officeParser.js> [--configOption=value] [FILE_PATH]
node officeparser [--configOption=value] [FILE_PATH]
```

Otherwise, you can simply use npx without installing the node module to instantly extract parsed data.
```
npx officeparser [--configOption=value] [FILE_PATH]
```

### Config Options:
- `--ignoreNotes=[true|false]`          Flag to ignore notes from files like PowerPoint. Default is false.
- `--newlineDelimiter=[delimiter]`      The delimiter to use for new lines. Default is `\n`.
- `--putNotesAtLast=[true|false]`       Flag to collect notes at the end of files like PowerPoint. Default is false.
- `--outputErrorToConsole=[true|false]` Flag to output errors to the console. Default is false.

## Library Usage
```js
const officeParser = require('officeparser');

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err) {
    // "data" string in the callback here is the text parsed from the office file passed in the first argument above
    if (err) {
        console.log(err);
        return;
    }
    console.log(data);
})

// promise
officeParser.parseOfficeAsync("/path/to/officeFile");
// "data" string in the promise here is the text parsed from the office file passed in the argument above
    .then(data => console.log(data))
    .catch(err => console.error(err))

// async/await
try {
    // "data" string returned from promise here is the text parsed from the office file passed in the argument
    const data = await officeParser.parseOfficeAsync("/path/to/officeFile");
    console.log(data);
} catch (err) {
    // resolve error
    console.log(err);
}

// USING FILE BUFFERS
// instead of file path, you can also pass file buffers of one of the supported files
// on parseOffice or parseOfficeAsync functions.

// get file buffers
const fileBuffers = fs.readFileSync("/path/to/officeFile");
// get parsed text from officeParser
// NOTE: Only works with parseOffice. Old functions are not supported.
officeParser.parseOfficeAsync(fileBuffers);
    .then(data => console.log(data))
    .catch(err => console.error(err))
```

### Configuration Object: OfficeParserConfig
*Optionally add a config object as 3rd variable to parseOffice for the following configurations*
| Flag                 | DataType | Default          | Explanation                                                                                                                                                                                                                                     |
|----------------------|----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| outputErrorToConsole | boolean  | false            | Flag to show all the logs to console in case of an error. Default is false.                                                                                                                                                                     |
| newlineDelimiter     | string   | \n               | The delimiter used for every new line in places that allow multiline text like word. Default is \n.                                                                                                                                             |
| ignoreNotes          | boolean  | false            | Flag to ignore notes from parsing in files like powerpoint. Default is false. It includes notes in the parsed text by default.                                                                                                                  |
| putNotesAtLast       | boolean  | false            | Flag, if set to true, will collectively put all the parsed text from notes at last in files like powerpoint. Default is false. It puts each notes right after its main slide content. If ignoreNotes is set to true, this flag is also ignored. |
<br>

```js
const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// callback
officeParser.parseOffice("/path/to/officeFile", function(data, err){
    if (err) {
        console.log(err);
        return;
    }
    console.log(data);
}, config)

// promise
officeParser.parseOfficeAsync("/path/to/officeFile", config);
    .then(data => console.log(data))
    .catch(err => console.error(err))
```

**Example - JavaScript**
```js
const officeParser = require('officeparser');

const config = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// relative path is also fine => eg: files/myWorkSheet.ods
officeParser.parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config);
    .then(data => {
        const newText = data + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm, filepath) {
    return officeParser.parseOfficeAsync(filepath)
        .then(data => data.indexOf(searchterm) != -1)
}
```


**Example - TypeScript**
```ts
import { OfficeParserConfig, parseOfficeAsync } from 'officeparser';

const config: OfficeParserConfig = {
    newlineDelimiter: " ",  // Separate new lines with a space instead of the default \n.
    ignoreNotes: true       // Ignore notes while parsing presentation files like pptx or odp.
}

// relative path is also fine => eg: files/myWorkSheet.ods
parseOfficeAsync("/Users/harsh/Desktop/files/mySlides.pptx", config);
    .then(data => {
        const newText = data + " look, I can parse a powerpoint file";
        callSomeOtherFunction(newText);
    })
    .catch(err => console.error(err));

// Search for a term in the parsed text.
function searchForTermInOfficeFile(searchterm: string, filepath: string): Promise<boolean> {
    return parseOfficeAsync(filepath)
        .then(data => data.indexOf(searchterm) != -1)
}
```
\
**Please take note: I have breached convention in placing err as second argument in my callback but please understand that I had to do it to not break other people's existing modules.**

## Browser Usage
Download the bundle file available as part of the release asset.
Include this bundle file in your browser html file and access `parseOffice` and `parseOfficeAsync` under the **`officeParser`** namespace.

**Example**
```html
<head>
    ...
    <!-- Include bundle file in the script tag. -->
    <script src="officeParserBundle@5.1.0.js"></script>
</head>
<body>
    ...
    <input type="file" id="fileInput" />
    ...
    <script>
        document.getElementById('fileInput').addEventListener('change', async function(event) {
            const outputDiv = document.getElementById('output');
            const file = event.target.files[0];
            try {
                // Your configuration options for officeParser
                const config = {
                    outputErrorToConsole: false,
                    newlineDelimiter: '\n',
                    ignoreNotes: false,
                    putNotesAtLast: false
                };

                const arrayBuffer = await file.arrayBuffer();
                const result = await officeParser.parseOfficeAsync(arrayBuffer, config);
                // result contains the extracted text.
            }
            catch (error) {
                // Handle error
            }
        });
    </script>
</body>
```


## Known Bugs
1. Inconsistency and incorrectness in the positioning of footnotes and endnotes in .docx files where the footnotes and endnotes would end up at the end of the parsed text whereas it would be positioned exactly after the referenced word in .odt files.
2. The charts and objects information of .odt files are not accurate and may end up showing a few NaN in some cases.
3. Extracting texts in browser bundles does not work for pdf files.
----------

**npm**
https://npmjs.com/package/officeparser

**github**
https://github.com/harshankur/officeParser
