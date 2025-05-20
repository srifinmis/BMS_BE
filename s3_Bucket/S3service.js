const express = require("express");
const AWS = require("aws-sdk");
const fs = require("fs");
const path = require("path");
const busboy = require("busboy");
// AWS S3 configuration
AWS.config.update({
  accessKeyId: process.env.AWS_ACCESS_KEY,
  secretAccessKey: process.env.AWS_SECRET_KEY,
  region: process.env.AWS_REGION,
});

const s3 = new AWS.S3();

const uploadFileToS3 = (req, res) => {
  const bb = busboy({ headers: req.headers });

  bb.on("field", (name, value) => {
    if (name === "filePath") {
      const filePath = value.trim();

      if (!fs.existsSync(filePath)) {
        return res.status(400).json({ error: "File does not exist on server" });
      }

      const filename = path.basename(filePath);
      const fileStream = fs.createReadStream(filePath);

      console.log(`Uploading file: ${filename}`);

      const params = {
        Bucket: process.env.AWS_BUCKET_NAME,
        Key: `uploads/${Date.now()}-${filename}`, // Unique filename
        Body: fileStream,
        ContentType: "application/pdf", // Adjust as needed
        ACL: "private",
      };

      s3.upload(params, (err, data) => {
        if (err) {
          console.error("S3 Upload Error:", err);
          return res.status(500).json({ error: "Error uploading to S3" });
        }

        console.log("File uploaded successfully:", data.Location);
        res.status(200).json({ s3Url: data.Location });
      });
    }
  });

  bb.on("error", (err) => {
    console.error("Busboy Error:", err);
    res.status(500).json({ error: "Error processing upload" });
  });

  req.pipe(bb);
};
module.exports = { uploadFileToS3 };
