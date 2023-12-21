const excelFileInput = document.getElementById("excel-file");
const searchButton = document.getElementById("search-button");
const studentNameInput = document.getElementById("student-name");
const resultsDiv = document.getElementById("results");
const fileName = document.getElementById("file-name");
const biology = document.getElementById("biology-mark");
const amharic = document.getElementById("amharic-mark");
const physics = document.getElementById("physics-mark");
const math = document.getElementById("math-mark");
const FoundstudentName = document.getElementById("name");

let studentData = [];
async function readFileAsync(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
}

window.addEventListener("keydown", (e) => {
  if (e.code == "Enter") {
    searchButton.click();
  }
});

excelFileInput.addEventListener("change", async (event) => {
  try {
    const file = event.target.files[0];

    if (
      (file.type ===
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" &&
        file.name.toLowerCase().endsWith(".xlsx")) ||
      file.name.toLowerCase().endsWith("xls")
    ) {
      console.log(file.name);
      fileName.textContent = file.name;

      const data = await readFileAsync(file);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];

      studentData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      console.log(studentData);
    } else {
      throw new Error(
        "Invalid file format. Please choose an Excel file (.xlsx)."
      );
    }
  } catch (error) {
    console.error("Error reading Excel file:", error);
    alert("Error reading Excel file. Please try again.");
  }
});

searchButton.addEventListener("click", () => {
  const studentName = studentNameInput.value.trim();
  console.log(studentName);

  if (!studentName) {
    alert("Please enter a valid student name.");
    return;
  }

  const foundStudent = studentData.find(
    (student) => student.Name.toLowerCase() == studentName.toLowerCase()
  );

  if (foundStudent) {
    FoundstudentName.textContent = foundStudent.Name;
    math.textContent = foundStudent.Math;
    physics.textContent = foundStudent.Physics;
    amharic.textContent = foundStudent.English;
    console.log(foundStudent);
  } else {
    alert("STUDENT NAME NOT FOUND");
  }
});
