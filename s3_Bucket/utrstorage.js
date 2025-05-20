const express = require("express");
const path = require("path");
const fs = require("fs");
const Busboy = require("busboy");
const router = express.Router();

const uploadDir = path.join(__dirname, "UTRuploads");

if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir, { recursive: true });
}

const UTRuploadFileToLocal = (req, res) => {
    const busboy = Busboy({ headers: req.headers });
    const filePaths = []; // Array to store file paths

    busboy.on("file", (fieldname, file, info) => {
        const filename = info.filename;
        if (!filename) {
            //  Don't return here, or you will exit before all files are processed.  Just skip.
            return;
        }

        const saveTo = path.join(uploadDir, filename);
        file.pipe(fs.createWriteStream(saveTo));

        file.on("end", () => {
            filePaths.push(saveTo); // Add path to the array
        });
    });

    busboy.on("finish", () => {
        if (filePaths.length > 0) {
            if (filePaths.length === 1) {
                res.json({ message: "Upload successful", filePath: filePaths[0] }); // send file path
            } else {
                res.json({ message: "Upload successful", filePaths: filePaths }); // Send array
            }

        } else {
            res.status(400).json({ error: "No files uploaded" });
        }
    });

    busboy.on("error", (err) => {
        console.error("Busboy Error:", err);
        res.status(500).json({ error: "File upload failed" });
    });

    req.pipe(busboy);
};

module.exports = { UTRuploadFileToLocal };