const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  // Class management
  getClasses: () => ipcRenderer.invoke("get-classes"),
  createClass: (data) => ipcRenderer.invoke("create-class", data),
  getTemplates: (className) => ipcRenderer.invoke("get-templates", className),
  uploadTemplate: (data) => ipcRenderer.invoke("upload-template", data),
  getClassSettings: (className) => ipcRenderer.invoke("get-class-settings", className),
  saveClassSettings: (data) => ipcRenderer.invoke("save-class-settings", data),
  
  // Student management
  getStudents: (className) => ipcRenderer.invoke("get-students", className),
  saveStudent: (data) => ipcRenderer.invoke("save-student", data),
  loadStudent: (data) => ipcRenderer.invoke("load-student", data),
  deleteStudent: (data) => ipcRenderer.invoke("delete-student", data),
  
  // Report generation
  generateStudentDocx: (data) => ipcRenderer.invoke("generate-student-docx", data),
  
  // Statistics
  getClassStats: (className) => ipcRenderer.invoke("get-class-stats", className),
  
  // File dialogs
  selectFile: (options) => ipcRenderer.invoke("select-file", options),
  selectFolder: () => ipcRenderer.invoke("select-folder"),
  openFolder: (folderPath) => ipcRenderer.invoke("open-folder", folderPath)
});