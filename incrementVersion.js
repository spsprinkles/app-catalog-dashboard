const fs = require("fs");
const path = require("path");

// Get the package file
const packageFile = require("./package.json");

// Increment the version #
const versionInfo = packageFile.version.toString().split('.');
versionInfo[2]++;
if (versionInfo[2] > 10) {
    versionInfo[1]++;
    versionInfo[2] -= 10;
}
if (versionInfo[1] > 10) {
    versionInfo[0]++;
    versionInfo[1] -= 10;
}
const version = versionInfo.join('.');

// Log
console.log("New Version: " + version);

// Update the package file
packageFile.version = version;
fs.writeFileSync("./package.json", JSON.stringify(packageFile, null, 2));
console.log("Package file updated");

// Get the strings.ts file
const stringsFile = fs.readFileSync("./src/strings.ts").toString().split('\n');
for (var i = 0; i < stringsFile.length; i++) {
    // See if this is the version #
    var versionIdx = stringsFile[i].indexOf("Version:");
    var startIdx = stringsFile[i].indexOf('"', versionIdx + 9);
    var endIdx = stringsFile[i].lastIndexOf('"');
    if (versionIdx > 0 && startIdx > versionIdx && endIdx > startIdx) {
        // Update the version
        stringsFile[i] = stringsFile[i].substring(0, startIdx + 1) + "0." + version + stringsFile[i].substring(endIdx);
        break;
    }
}
fs.writeFileSync("./src/strings.ts", stringsFile.join('\n'));
console.log("Strings file updated");

// Get the SPFx package-solution file
const packageSolutionFile = require("./spfx/config/package-solution.json");
packageSolutionFile.solution.version = "0." + version;

// Update the SPFx package-solution file
fs.writeFileSync("./spfx/config/package-solution.json", JSON.stringify(packageSolutionFile, null, 2));
console.log("SPFx package-solution file updated");
