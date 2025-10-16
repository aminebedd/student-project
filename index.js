import express from "express";
import mongoose from "mongoose";
import ExcelJS from "exceljs";
import cors from "cors";

const app = express();
app.use(cors()); 
app.use(express.json()); 
const studentSchema = new mongoose.Schema({
  firstname: String,
  lastname: String,
  email: String,
  phonenumber: String,
  speciality: String,
});

const Student = mongoose.model("Student", studentSchema);

// ✅ إضافة طالب جديد
app.post("/add-student", async (req, res) => {
  try {
    const { firstname, lastname, email, phonenumber, speciality } = req.body;

    if (!firstname || !lastname || !email) {
      return res.status(400).json({ message: "⚠️ Missing required fields" });
    }

    const existingStudent = await Student.findOne({ email });
    if (existingStudent) {
      return res.status(400).json({ message: "❌ Email already exists" });
    }

    const newStudent = new Student({
      firstname,
      lastname,
      email,
      phonenumber,
      speciality,
    });

    await newStudent.save();
    res.status(201).json({
      message: "✅ Student added successfully",
      student: newStudent,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ message: "❌ Error adding student", error: err.message });
  }
});

// ✅ تحميل ملف Excel
app.get("/download-excel", async (req, res) => {
  try {
    const students = await Student.find();

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Students");

    worksheet.columns = [
      { header: "First Name", key: "firstname", width: 20 },
      { header: "Last Name", key: "lastname", width: 20 },
      { header: "Email", key: "email", width: 30 },
      { header: "Phone number", key: "phonenumber", width: 15 },
      { header: "Speciality", key: "speciality", width: 25 },
    ];

    students.forEach((student) => worksheet.addRow(student));

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=students.xlsx");

    await workbook.xlsx.write(res);
    res.status(200).end();
  } catch (err) {
    console.error(err);
    res.status(500).send("Error generating Excel file");
  }
});

// ✅ اتصال قاعدة البيانات
mongoose
  .connect("mongodb+srv://mbeddani00:19731976moh@cluster0.hapjev4.mongodb.net/Studentmi")
  .then(() => {
    app.listen(3000, () => console.log("✅ Server running on port 3000"));
  })
  .catch((err) => console.error(err));
