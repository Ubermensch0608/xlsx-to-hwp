import express, { NextFunction } from "express";
import path from "path";
import multer from "multer";

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
  console.log(req.files, "req.files");
  console.log(req.body, "req.body");

  res.json(req.files);
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
  console.log(`http://localhost:3000`);
});
