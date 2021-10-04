var path = require("path");
var project = require("./package.json");

// Export the configuration
module.exports = (env, argv) => {
    var isDev = argv.mode === "development";

    // Return the configuration
    return {
        // Main project files
        entry: [
            //"./node_modules/gd-sprest-bs/dist/gd-sprest-bs.min.js",
            path.resolve(__dirname, project.main)
        ],

        // Output information
        output: {
            path: path.resolve(__dirname, "dist"),
            filename: project.name + (isDev ? "" : ".min") + ".js"
        },

        // External Libs
        // Since we are including the library as part of the entry point, we don't
        // want webpack to bundle the library twice.
        /*externals: {
            "gd-sprest": "$REST",
            "gd-sprest-bs": "$REST"
        },*/

        // Resolve the file names
        resolve: {
            extensions: [".js", ".ts"]
        },

        // Compiler Information
        module: {
            rules: [
                // Handle CSS Files
                {
                    test: /\.css$/,
                    use: [
                        // Create the style nodes from the CommonJS code
                        { loader: "style-loader" },
                        // Translate css to CommonJS
                        { loader: "css-loader" }
                    ]
                },
                // Handle HTML Files
                {
                    test: /\.html$/,
                    exclude: "/node_modules/",
                    use: [{ loader: "html-loader" }]
                },
                // Handle Image Files
                {
                    test: /\.(jpe?g|png|gif|svg)$/,
                    use: "url-loader"
                },
                // Handle TypeScript Files
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: [
                        // Step 2 - Compile JavaScript ES6 to JavaScript Current Standards
                        {
                            loader: "babel-loader",
                            options: {
                                presets: ["@babel/preset-env"]
                            }
                        },
                        // Step 1 - Compile TypeScript to JavaScript ES6
                        {
                            loader: "ts-loader"
                        }
                    ]
                }
            ]
        }
    };
}