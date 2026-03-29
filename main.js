const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const { autoUpdater } = require("electron-updater");

// Fix ICU data error on Windows (Electron 38+)
app.commandLine.appendSwitch('disable-features', 'IcuInitializeEnforced');
app.commandLine.appendSwitch('no-sandbox');

// Helper: resolve writable user data path (works both in dev and packaged)
function userDataPath(...segments) {
  return path.join(app.getPath('userData'), ...segments);
}

// Configure auto-updater to run asynchronously (don't block startup)
setTimeout(() => {
  autoUpdater.checkForUpdatesAndNotify();
}, 3000); // Check after 3 seconds

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  win.loadFile(path.join(__dirname, "renderer/index.html"));
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

// ==================== DATA MANAGEMENT ====================

// Directory structure: classes/{classname}/templates/{template}.docx
// Directory structure: classes/{classname}/students/{studentname}.json

// Create directories if they don't exist
function ensureDirectories() {
  const baseDir = userDataPath("classes");
  if (!fs.existsSync(baseDir)) {
    fs.mkdirSync(baseDir, { recursive: true });
  }
}

// Get all classes
ipcMain.handle("get-classes", async () => {
  ensureDirectories();
  const classesDir = userDataPath("classes");
  if (!fs.existsSync(classesDir)) return [];

  const classes = fs.readdirSync(classesDir, { withFileTypes: true })
    .filter(dirent => dirent.isDirectory())
    .map(dirent => dirent.name);

  return classes;
});

// Create a new class with template
ipcMain.handle("create-class", async (event, { className, templatePath }) => {
  ensureDirectories();
  const classDir = userDataPath("classes", className);
  const templatesDir = path.join(classDir, "templates");
  const studentsDir = path.join(classDir, "students");

  // Create directories
  fs.mkdirSync(templatesDir, { recursive: true });
  fs.mkdirSync(studentsDir, { recursive: true });

  // Copy template if provided
  if (templatePath && fs.existsSync(templatePath)) {
    const templateName = path.basename(templatePath);
    const destPath = path.join(templatesDir, templateName);
    fs.copyFileSync(templatePath, destPath);
  }

  return { success: true, classDir };
});

// Get templates for a class
ipcMain.handle("get-templates", async (event, className) => {
  const templatesDir = userDataPath("classes", className, "templates");
  if (!fs.existsSync(templatesDir)) return [];

  const templates = fs.readdirSync(templatesDir)
    .filter(file => file.endsWith('.docx'))
    .map(file => ({
      name: path.basename(file, '.docx'),
      path: path.join(templatesDir, file)
    }));

  return templates;
});

// Upload a template to a class
ipcMain.handle("upload-template", async (event, { className, filePath, sourcePath }) => {
  const src = filePath || sourcePath;  // accept both property names
  if (!src) return { success: false, error: 'No file path provided' };

  const templatesDir = userDataPath("classes", className, "templates");
  if (!fs.existsSync(templatesDir)) {
    fs.mkdirSync(templatesDir, { recursive: true });
  }

  const templateName = path.basename(src);
  const destPath = path.join(templatesDir, templateName);
  fs.copyFileSync(src, destPath);

  return { success: true, templatePath: destPath, templateName: path.basename(templateName, '.docx') };
});

// Get students in a class
ipcMain.handle("get-students", async (event, className) => {
  const studentsDir = userDataPath("classes", className, "students");
  if (!fs.existsSync(studentsDir)) return [];

  const students = fs.readdirSync(studentsDir)
    .filter(file => file.endsWith('.json'))
    .map(file => {
      const filePath = path.join(studentsDir, file);
      let displayName = path.basename(file, '.json'); // fallback
      let isDraft = true;
      try {
        const saved = JSON.parse(fs.readFileSync(filePath, 'utf8'));
        if (saved.studentName) displayName = saved.studentName;
        isDraft = saved.isDraft !== false;
      } catch(e) { /* ignore corrupt files */ }
      return {
        name: displayName,
        fileKey: path.basename(file, '.json'), // the sanitized filename key for loading
        path: filePath,
        isDraft,
        lastModified: fs.statSync(filePath).mtime
      };
    });

  return students;
});


// Save student progress
ipcMain.handle("save-student", async (event, { className, studentName, data, isDraft = true }) => {
  ensureDirectories();
  const studentsDir = userDataPath("classes", className, "students");

  if (!fs.existsSync(studentsDir)) {
    fs.mkdirSync(studentsDir, { recursive: true });
  }

  const cleanName = studentName.replace(/[^a-zA-Z0-9_\s-]/g, '').replace(/\s+/g, '_');
  const fileName = `${cleanName}.json`;
  const filePath = path.join(studentsDir, fileName);

  const saveData = {
    ...data,
    savedAt: new Date().toISOString(),
    isDraft,
    className,
    studentName
  };

  fs.writeFileSync(filePath, JSON.stringify(saveData, null, 2));

  return {
    success: true,
    message: isDraft ? "Progress saved as draft" : "Student data saved",
    filePath
  };
});

// Load student data
ipcMain.handle("load-student", async (event, { className, studentName }) => {
  const cleanName = studentName.replace(/[^a-zA-Z0-9_\s-]/g, '').replace(/\s+/g, '_');
  const filePath = userDataPath("classes", className, "students", `${cleanName}.json`);

  if (!fs.existsSync(filePath)) {
    return { success: false, error: "Student data not found" };
  }

  const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  return { success: true, data };
});

// Generate DOCX for a student
ipcMain.handle("generate-student-docx", async (event, { className, studentName, templateName, data }) => {
  try {
    // Get the template path
    const templatePath = userDataPath("classes", className, "templates", `${templateName}.docx`);

    if (!fs.existsSync(templatePath)) {
      throw new Error(`Template "${templateName}" not found for class "${className}"`);
    }

    const content = fs.readFileSync(templatePath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip);

    // Calculate grades and remarks
    const subjectsWithDetails = data.subjects.map((s, index) => {
      const score = s.score || 0;
      let grade = "", remark = "";

      if (score >= 70) { grade = "A"; remark = "Excellent"; }
      else if (score >= 60) { grade = "B"; remark = "Very Good"; }
      else if (score >= 50) { grade = "C"; remark = "Good"; }
      else if (score >= 40) { grade = "D"; remark = "Pass"; }
      else { grade = "E"; remark = "Fail"; }

      return {
        index: index + 1,
        subject: s.subject,
        test1: s.test1 || 0,      // ⬅️ ADD THIS
        test2: s.test2 || 0,      // ⬅️ ADD THIS
        exam: s.exam || 0,
        score: score,
        grade: grade,
        remark: remark
      };
    });

    // Calculate summary
    const totalSubjects = data.subjects.length;
    let passCount = 0;
    let failCount = 0;
    let totalScore = 0;

    data.subjects.forEach(s => {
      const score = s.score || 0;
      totalScore += score;
      if (score >= 40) passCount++;
      else failCount++;
    });

    const average = totalSubjects > 0 ? (totalScore / totalSubjects).toFixed(2) : 0;

    // Prepare template data
    const templateData = {
      studentName: data.studentName || 'N/A',
      studentClass: data.studentClass || 'N/A',
      term: data.term || 'N/A',
      academicYear: data.academicYear || 'N/A',
      numberInClass: data.numberInClass || 'N/A',
      classHighestScore: data.classHighestScore || 'N/A',
      nextTermBegins: data.nextTermBegins || 'N/A',

      subjects: subjectsWithDetails,

      // Psychomotor skills
      psy_handwriting: data.psychomotor?.Handwriting || 'N/A',
      psy_fluency: data.psychomotor?.Fluency || 'N/A',
      psy_games: data.psychomotor?.Games || 'N/A',
      psy_craft: data.psychomotor?.Craft || 'N/A',
      psy_drawing: data.psychomotor?.Drawing || 'N/A',
      psy_handling: data.psychomotor?.["Handling Tools"] || 'N/A',

      // Affective disposition
      aff_punctuality: data.affective?.Punctuality || 'N/A',
      aff_attendance: data.affective?.Attendance || 'N/A',
      aff_reliability: data.affective?.Reliability || 'N/A',
      aff_neatness: data.affective?.Neatness || 'N/A',
      aff_politeness: data.affective?.Politeness || 'N/A',
      aff_honesty: data.affective?.Honesty || 'N/A',
      aff_staff: data.affective?.["Staff Relation"] || 'N/A',
      aff_students: data.affective?.["Student Relation"] || 'N/A',
      aff_selfcontrol: data.affective?.["Self Control"] || 'N/A',
      aff_responsibility: data.affective?.Responsibility || 'N/A',
      aff_attentiveness: data.affective?.Attentiveness || 'N/A',
      aff_spiritual: data.affective?.["Spiritual Cooperation"] || 'N/A',
      aff_organization: data.affective?.Organization || 'N/A',
      aff_perseverance: data.affective?.Perseverance || 'N/A',


      // Summary
      totalSubjects: totalSubjects,
      passCount: passCount,
      failCount: failCount,
      totalScore: totalScore,
      average: average,
      date: new Date().toLocaleDateString()
    };

    doc.render(templateData);

    // Generate the DOCX file
    const buf = doc.getZip().generate({ type: "nodebuffer" });

    // Save the generated report
    const reportsDir = userDataPath("classes", className, "reports");
    if (!fs.existsSync(reportsDir)) {
      fs.mkdirSync(reportsDir, { recursive: true });
    }

    const cleanName = (studentName || "Report").replace(/[^a-zA-Z0-9_\s-]/g, '').replace(/\s+/g, '_');
    const fileName = `${cleanName}_${data.term || ''}.docx`;
    const filePath = path.join(reportsDir, fileName);

    fs.writeFileSync(filePath, buf);

    // Save/update student data to mark as completed
    const studentsDir = userDataPath("classes", className, "students");
    if (!fs.existsSync(studentsDir)) {
      fs.mkdirSync(studentsDir, { recursive: true });
    }
    const studentFilePath = path.join(studentsDir, `${cleanName}.json`);

    let studentData = {};
    if (fs.existsSync(studentFilePath)) {
      studentData = JSON.parse(fs.readFileSync(studentFilePath, 'utf8'));
    } else {
      // Create new student data if file doesn't exist
      studentData = data;
    }

    // Update with generation metadata
    studentData.generatedAt = new Date().toISOString();
    studentData.generatedFile = filePath;
    studentData.isDraft = false;
    studentData.className = className;
    studentData.studentName = data.studentName;
    studentData.savedAt = new Date().toISOString();

    fs.writeFileSync(studentFilePath, JSON.stringify(studentData, null, 2));

    // Show success notification
    const win = BrowserWindow.getFocusedWindow();
    if (win) {
      dialog.showMessageBox(win, {
        type: 'info',
        title: 'DOCX Generated',
        message: `Report generated for ${studentName}`,
        detail: `Saved to: ${filePath}`,
        buttons: ['OK', 'Open Folder']
      }).then(result => {
        if (result.response === 1) {
          const { shell } = require('electron');
          shell.openPath(reportsDir);
        }
      });
    }

    return {
      success: true,
      filePath: filePath,
      fileName: fileName
    };

  } catch (err) {
    console.error("DOCX Export Error:", err);

    const win = BrowserWindow.getFocusedWindow();
    if (win) {
      dialog.showErrorBox('DOCX Generation Failed', err.message);
    }

    return {
      success: false,
      error: err.message
    };
  }
});

// Get student statistics
ipcMain.handle("get-class-stats", async (event, className) => {
  const studentsDir = userDataPath("classes", className, "students");
  const reportsDir = userDataPath("classes", className, "reports");

  if (!fs.existsSync(studentsDir)) {
    return { totalStudents: 0, completed: 0, pending: 0 };
  }

  const students = fs.readdirSync(studentsDir).filter(file => file.endsWith('.json'));
  const reports = fs.existsSync(reportsDir) ?
    fs.readdirSync(reportsDir).filter(file => file.endsWith('.docx')) : [];

  let completed = 0;
  let pending = 0;

  students.forEach(studentFile => {
    const studentPath = path.join(studentsDir, studentFile);
    const studentData = JSON.parse(fs.readFileSync(studentPath, 'utf8'));
    if (studentData.generatedAt && !studentData.isDraft) {
      completed++;
    } else {
      pending++;
    }
  });

  return {
    totalStudents: students.length,
    completed,
    pending,
    totalReports: reports.length
  };
});

// Delete student data
ipcMain.handle("delete-student", async (event, { className, studentName }) => {
  const cleanName = studentName.replace(/[^a-zA-Z0-9_\s-]/g, '').replace(/\s+/g, '_');
  const studentPath = userDataPath("classes", className, "students", `${cleanName}.json`);

  if (fs.existsSync(studentPath)) {
    fs.unlinkSync(studentPath);
    return { success: true, message: "Student data deleted" };
  }

  return { success: false, error: "Student not found" };
});

// Select folder for custom data location
ipcMain.handle("select-folder", async (event) => {
  const win = BrowserWindow.getFocusedWindow();
  if (!win) return null;

  try {
    const result = await dialog.showOpenDialog(win, {
      properties: ['openDirectory', 'createDirectory'],
      message: 'Select folder to store classes and student data'
    });

    if (!result.canceled && result.filePaths.length > 0) {
      return result.filePaths[0];
    }
    return null;
  } catch (error) {
    console.error("Folder selection error:", error);
    return null;
  }
});

// Input dialog for class creation
ipcMain.handle("show-class-input", async (event) => {
  const win = BrowserWindow.getFocusedWindow();
  if (!win) return null;

  try {
    const { response: input } = await dialog.showMessageBox(win, {
      type: 'question',
      title: 'Create New Class',
      message: 'Enter the class name (e.g., JSS 1A, Primary 5):',
      detail: 'This will create a new folder to store students and reports for this class.',
      buttons: ['Cancel', 'Create'],
      defaultId: 1,
      cancelId: 0
    });

    if (input === 1) {
      // Show an input-like dialog - actually use a workaround
      // For now, return request to frontend to use browser prompt
      return 'use-browser-prompt';
    }
    return null;
  } catch (error) {
    console.error("Input dialog error:", error);
    return null;
  }
});

// ==================== FILE DIALOGS ====================

// Open a native file-picker and return the selected path
ipcMain.handle("select-file", async (event, options = {}) => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win, {
    title: options.title || "Select File",
    filters: options.filters || [{ name: "All Files", extensions: ["*"] }],
    properties: ["openFile"]
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

// Open a folder in Windows Explorer
ipcMain.handle("open-folder", async (event, folderPath) => {
  const { shell } = require("electron");
  if (folderPath && fs.existsSync(folderPath)) {
    await shell.openPath(folderPath);
    return { success: true };
  }
  return { success: false, error: "Folder not found" };
});

// Get class-level settings
ipcMain.handle("get-class-settings", async (event, className) => {
  const settingsPath = userDataPath("classes", className, "settings.json");
  if (fs.existsSync(settingsPath)) {
    try {
      const data = fs.readFileSync(settingsPath, 'utf8');
      return JSON.parse(data);
    } catch (err) {
      return { academicYear: "", term: "1st Term", nextTermBegins: "", classHighestScore: "" };
    }
  }
  return { academicYear: "", term: "1st Term", nextTermBegins: "", classHighestScore: "" };
});

// Save class-level settings
ipcMain.handle("save-class-settings", async (event, { className, settings }) => {
  try {
    const classDir = userDataPath("classes", className);
    if (!fs.existsSync(classDir)) {
      fs.mkdirSync(classDir, { recursive: true });
    }
    const settingsPath = path.join(classDir, "settings.json");
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2));
    return { success: true };
  } catch (err) {
    console.error("Error saving class settings:", err);
    return { success: false, error: err.message };
  }
});