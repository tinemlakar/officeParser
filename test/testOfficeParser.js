// @ts-check

const officeParser = require("../officeParser");
const fs = require("fs");
const supportedExtensions = require("../supportedExtensions");

// File names of test files and their text output content
// test file name style => test.<ext>
// test content output => test.<ext>.txt

/** List of all supported extensions with office Parser */
const supportedExtensionTests = [
  {
    ext: "docx",
    testAvailable: true,
  },
  {
    ext: "xlsx",
    testAvailable: true,
  },
  {
    ext: "pptx",
    testAvailable: true,
  },
  {
    ext: "odt",
    testAvailable: true,
  },
  {
    ext: "odp",
    testAvailable: true,
  },
  {
    ext: "ods",
    testAvailable: true,
  },
];

/** Config file for performing tests */
const config = {
  preserveTempFiles: true,
  outputErrorToConsole: true,
};

/** Local list of supported extensions in test file */
const localSupportedExtensionsList = supportedExtensionTests.map(
  (test) => test.ext
);

/** Get filename for an extension */
function getFilename(ext, isContentFile = false) {
  return `test/files/test.${ext}` + (isContentFile ? `.txt` : "");
}

/** Run test for a passed extension */
function runTest(ext, buffer) {
  return officeParser
    .parseOfficeAsync(
      buffer ? fs.readFileSync(getFilename(ext)) : getFilename(ext),
      config
    )
    .then((text) =>
      fs.readFileSync(getFilename(ext, true), "utf8") == text
        ? console.log(
            `[${ext.padEnd(4)}: ${buffer ? "buffer" : "file  "}] => Passed`
          )
        : console.log(
            `[${ext.padEnd(4)}: ${buffer ? "buffer" : "file  "}] => Failed`
          )
    )
    .catch((error) => console.log("ERROR: " + error));
}

async function runAllTests() {
  for (let i = 0; i < supportedExtensionTests.length; i++) {
    const test = supportedExtensionTests[i];
    if (test.testAvailable) {
      await runTest(test.ext, false);
      await runTest(test.ext, true);
    } else console.log(`[${test.ext}]=> Skipped`);
  }
}

// Run all test files with test content if no argument passed.
if (process.argv.length == 2) {
  // Test to check all items in local extension list are present in supportedExtensions.js file
  localSupportedExtensionsList.every((ext) => supportedExtensions.includes(ext))
    ? console.log(
        "All extensions in test files found in primary supportedExtensions.js file"
      )
    : console.warn(
        "Extension in test files missing from primary supportedExtensions.js file"
      );

  // Test to check all items in supportedExtensions.js file are present in local extension list
  supportedExtensions.every((ext) => localSupportedExtensionsList.includes(ext))
    ? console.log(
        "All extensions in primary supportedExtensions.js file found in test file"
      )
    : console.warn(
        "Extension in primary supportedExtensions.js file missing from test file"
      );

  runAllTests();
} else if (process.argv.length == 3) {
  if (localSupportedExtensionsList.includes(process.argv[2]))
    officeParser
      .parseOfficeAsync(getFilename(process.argv[2]), config)
      .then((text) => console.log(text))
      .catch((error) => console.log("ERROR: " + error));
  else
    console.error("The requested extension test is not currently available.");
} else console.error("Invalid arguments");
