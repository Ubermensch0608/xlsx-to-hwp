import express, { NextFunction } from "express";
import path from "path";
import multer from "multer";
import xlsx from "xlsx";
import fs from "fs";

const app = express();

// middleware
app.use(express.static(path.join(__dirname, "../public")));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// multer
const upload = multer({ dest: "uploads/" });

//500 에러 핸들링
app.use((err: Error, _req: unknown, res: any, next: NextFunction) => {
  console.error(err.stack);
  res.status(500).send("Something broke!");
  next(err);
});

// routes
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "../public/index.html"));
});

app.post("/convert", upload.any(), (req, res) => {
  if (!req.files) {
    return res.status(400).json({ message: "파일이 없습니다." });
  }

  const excelFile = Object.values(req.files).find(
    (file) => file.fieldname === "xlsxFile"
  );

  const targetExcel = xlsx.readFile(excelFile.path);
  const sheetName = targetExcel.SheetNames[0]; // 첫 번째 시트 기준
  const worksheet = targetExcel.Sheets[sheetName];
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

  res.json(jsonData);

  if (req.files) {
    for (const file of Object.values(req.files)) {
      // 임시 파일 삭제
      fs.unlinkSync(file.path);
    }
  }
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
  console.log(`http://localhost:3000`);
});
