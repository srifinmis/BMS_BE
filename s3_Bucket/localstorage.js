// const express = require("express");
// const multer = require("multer");
// const path = require("path");
// const fs = require("fs");


// // Ensure 'uploads' folder exists
// const uploadDir = path.join(__dirname, "uploads");
// if (!fs.existsSync(uploadDir)) {
//   fs.mkdirSync(uploadDir, { recursive: true });
// }

// // Multer configuration for storing files
// const storage = multer.diskStorage({
//   destination: (req, file, cb) => {
//     cb(null, uploadDir); // Save files in 'uploads' folder
//   },
//   filename: (req, file, cb) => {
//     const uniqueName = `${Date.now()}-${Math.round(Math.random() * 1E9)}${path.extname(file.originalname)}`;
//     cb(null, uniqueName); // Ensure unique file names
//   },
// });

// const upload = multer({ 
//   storage,
//   limits: { fileSize: 10 * 1024 * 1024 }, // Limit file size to 10MB
//   fileFilter: (req, file, cb) => {
//     const allowedTypes = ["image/jpeg", "image/png", "application/pdf"];
//     if (!allowedTypes.includes(file.mimetype)) {
//       return cb(new Error("Only JPEG, PNG, and PDF files are allowed"), false);
//     }
//     cb(null, true);
//   },
// });

// // Route to handle file upload
// const uploadFileToLocal = async (req, res) => {
//   console.log("Uploaded File:", req.file); // Debugging log

//   if (!req.file) {
//     return res.status(400).json({ error: "No file uploaded" });
//   }

//   const filePath = `./uploads/${req.file.name}`;

//   return res.json({
//     message: "File uploaded successfully!",
//     filePath,
//   });
// };

// module.exports = { uploadFileToLocal,upload };





const express = require("express");
const path = require("path");
const fs = require("fs");
// const router = express.Router();
const Busboy = require("busboy");
console.log(require("busboy"));
const router = express.Router();
const uploadDir = path.join(__dirname, "uploads");

if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const uploadFileToLocal = (req, res) => {
     const busboy = Busboy({ headers: req.headers }); 

  busboy.on("file", (fieldname, file, info) => {
    const filename = info.filename;
    if (!filename) {
      return res.status(400).json({ error: "No file uploaded" });
    }

    const saveTo = path.join(uploadDir, filename);
    file.pipe(fs.createWriteStream(saveTo));

    file.on("end", () => {
      res.json({ message: "Upload successful", filePath: saveTo });
    });
  });

  busboy.on("error", (err) => {
    console.error("Busboy Error:", err);
    res.status(500).json({ error: "File upload failed" });
  });

  req.pipe(busboy); // ✅ Must pipe request to busboy
};


module.exports = {uploadFileToLocal};