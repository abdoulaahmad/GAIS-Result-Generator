const express = require("express");
const fs = require("fs");
const cors = require("cors");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");


const app = express();
app.use(express.json());
app.use(cors());

app.post("/generate-report", (req, res) => {
  const { studentName, studentClass, term, subjects } = req.body;

  const template = fs.readFileSync("./SSS Exams Results.docx", "binary");
  const zip = new PizZip(template);
  const doc = new Docxtemplater(zip);

  doc.setData({
    studentName,
    studentClass,
    term,
    subjects,
  });

  try {
    doc.render();
    const buf = doc.getZip().generate({ type: "nodebuffer" });
    const outputPath = `./reports/${studentName}_report.docx`;
    fs.writeFileSync(outputPath, buf);
    res.json({ success: true, path: outputPath });
  } catch (error) {
    console.error(error);
    res.status(500).json({ success: false, error: error.message });
  }
});

app.listen(5000, () => console.log("Server running on port 5000"));
