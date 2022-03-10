const path = require("path");
const fs = require("fs");

// Get the image file
const imgFile = path.join(__dirname, "assets/AppIcon-96.png");

// See if it exists
if (fs.existsSync(imgFile)) {
    // Convert the image to base64
    var imgString = fs.readFileSync(imgFile, "base64");

    // Log
    console.log("data:image/png;base64," + imgString);
} else {
    // Log
    console.log("Icon file doesn't exist");
}