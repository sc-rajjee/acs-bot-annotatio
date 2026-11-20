const path = require("path");

module.exports = {

    mode: "development",

    entry: "./client.js",

    output: {
        filename: "bundle.js",
        path: path.resolve(
            __dirname,
            "public"
        )
    },

    devtool: "inline-source-map",

    devServer: {
        static: {
            directory: path.join(
                __dirname,
                "public"
            )
        },

        port: 8080,

        hot: true
    }
};