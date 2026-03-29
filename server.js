const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const session = require("express-session");
const XLSX = require("xlsx");
const PDFDocument = require("pdfkit");

const app = express();
const PORT = 3000;
const ADMIN_USERNAME = process.env.ADMIN_USERNAME || "admin";
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || "event123";

const dataDir = path.join(__dirname, "data");
const uploadsDir = path.join(__dirname, "uploads");
const dbFile = path.join(dataDir, "registrations.json");
const settingsFile = path.join(dataDir, "settings.json");

const defaultSettings = {
  registrationOpen: true,
  closedMessage: "REGISTRATION IS CLOSED.",
};

fs.mkdirSync(dataDir, { recursive: true });
fs.mkdirSync(uploadsDir, { recursive: true });
if (!fs.existsSync(dbFile)) {
  fs.writeFileSync(dbFile, "[]", "utf-8");
}
if (!fs.existsSync(settingsFile)) {
  fs.writeFileSync(
    settingsFile,
    JSON.stringify(defaultSettings, null, 2),
    "utf-8",
  );
}

function readSettings() {
  try {
    const raw = fs.readFileSync(settingsFile, "utf-8");
    return { ...defaultSettings, ...JSON.parse(raw) };
  } catch {
    return { ...defaultSettings };
  }
}

function writeSettings(next) {
  fs.writeFileSync(settingsFile, JSON.stringify(next, null, 2), "utf-8");
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => {
    const safeName = file.originalname.replace(/\s+/g, "_");
    cb(null, `${Date.now()}-${safeName}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ["image/png", "image/jpeg", "image/webp"];
    if (!allowed.includes(file.mimetype)) {
      return cb(new Error("Only PNG, JPEG, WEBP files are allowed"));
    }
    cb(null, true);
  },
});

app.use(express.json());
app.use(
  session({
    secret: process.env.SESSION_SECRET || "replace-with-strong-session-secret",
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      maxAge: 1000 * 60 * 60 * 4,
    },
  }),
);
app.use("/uploads", express.static(uploadsDir));

app.get("/admin.html", (req, res) => {
  if (req.session && req.session.isAdmin) {
    return res.sendFile(path.join(__dirname, "admin.html"));
  }
  return res.redirect("/admin");
});

app.use(express.static(__dirname));

function readRegistrations() {
  const raw = fs.readFileSync(dbFile, "utf-8");
  return JSON.parse(raw);
}

function requireAdmin(req, res, next) {
  if (req.session && req.session.isAdmin) {
    return next();
  }
  return res.status(401).json({ message: "Unauthorized" });
}

function toFlatRows(records) {
  return records.map((record) => ({
    id: record.id,
    createdAt: record.createdAt,
    email: record.email,
    teamName: record.teamName,
    teamPhone: record.teamPhone,
    teamCollege: record.teamCollege,
    domain: record.domain,
    teamSize: record.teamSize,
    transactionId: record.transactionId,
    paymentScreenshot: record.paymentScreenshot,
    members: (record.members || [])
      .map((m) => `${m.role}: ${m.name} (${m.usn}, ${m.phone})`)
      .join(" | "),
  }));
}

app.get("/admin", (req, res) => {
  if (req.session && req.session.isAdmin) {
    return res.sendFile(path.join(__dirname, "admin.html"));
  }
  return res.sendFile(path.join(__dirname, "admin-login.html"));
});

app.post("/api/admin/login", (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_USERNAME && password === ADMIN_PASSWORD) {
    req.session.isAdmin = true;
    return res.json({ message: "Login successful" });
  }
  return res.status(401).json({ message: "Invalid credentials" });
});

app.post("/api/admin/logout", (req, res) => {
  req.session.destroy(() => {
    res.json({ message: "Logged out" });
  });
});

app.get("/api/registration-status", (req, res) => {
  const s = readSettings();
  return res.json({
    open: Boolean(s.registrationOpen),
    message: String(s.closedMessage || defaultSettings.closedMessage),
  });
});

app.get("/api/admin/settings", requireAdmin, (req, res) => {
  return res.json({ data: readSettings() });
});

app.post("/api/admin/settings/registration", requireAdmin, (req, res) => {
  const { open, closedMessage } = req.body;
  const current = readSettings();
  const next = {
    ...current,
    registrationOpen:
      typeof open === "boolean" ? open : Boolean(current.registrationOpen),
  };
  if (typeof closedMessage === "string" && closedMessage.trim() !== "") {
    next.closedMessage = closedMessage.trim();
  }
  writeSettings(next);
  return res.json({ data: next });
});

app.get("/api/admin/registrations", requireAdmin, (req, res) => {
  try {
    const records = readRegistrations();
    return res.json({ data: records });
  } catch (error) {
    return res.status(500).json({ message: "Failed to fetch records" });
  }
});

app.get("/api/admin/download/excel", requireAdmin, (req, res) => {
  try {
    const records = readRegistrations();
    const rows = toFlatRows(records);
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Registrations");
    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="registrations-${Date.now()}.xlsx"`,
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    return res.send(buffer);
  } catch (error) {
    return res.status(500).json({ message: "Failed to generate Excel file" });
  }
});

app.get("/api/admin/download/pdf", requireAdmin, (req, res) => {
  try {
    const records = readRegistrations();
    const rows = toFlatRows(records);

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="registrations-${Date.now()}.pdf"`,
    );
    res.setHeader("Content-Type", "application/pdf");

    const doc = new PDFDocument({ margin: 30, size: "A4" });
    doc.pipe(res);
    doc.fontSize(16).text("Event Registrations", { align: "center" });
    doc.moveDown();
    doc.fontSize(10);

    rows.forEach((row, index) => {
      doc.text(`Entry ${index + 1}`, { underline: true });
      doc.text(`Date: ${row.createdAt}`);
      doc.text(
        `Team: ${row.teamName} | Domain: ${row.domain} | Size: ${row.teamSize}`,
      );
      doc.text(`Email: ${row.email} | Team Phone: ${row.teamPhone}`);
      doc.text(`College: ${row.teamCollege}`);
      doc.text(`Transaction: ${row.transactionId}`);
      doc.text(`Members: ${row.members}`);
      doc.text(`Screenshot: ${row.paymentScreenshot}`);
      doc.moveDown();
      if (doc.y > 760) {
        doc.addPage();
      }
    });

    doc.end();
  } catch (error) {
    return res.status(500).json({ message: "Failed to generate PDF file" });
  }
});

app.post("/api/register", upload.single("paymentScreenshot"), (req, res) => {
  try {
    const settings = readSettings();
    if (!settings.registrationOpen) {
      if (req.file && req.file.path) {
        try {
          fs.unlinkSync(req.file.path);
        } catch {}
      }
      return res.status(403).json({
        message: settings.closedMessage || defaultSettings.closedMessage,
      });
    }

    const {
      email,
      teamName,
      teamPhone,
      domain,
      teamSize,
      transactionId,
      teamCollege,
    } = req.body;

    const size = Number(teamSize);
    if (!(size === 3 || size === 4)) {
      return res.status(400).json({ message: "Invalid team size" });
    }

    const members = [];
    for (let i = 1; i <= size; i += 1) {
      const memberName = req.body[`member${i}Name`];
      const memberUsn = req.body[`member${i}Usn`];
      const memberPhone = req.body[`member${i}Phone`];
      if (!memberName || !memberUsn || !memberPhone) {
        return res.status(400).json({ message: "Member details missing" });
      }
      members.push({
        role: i === 1 ? "Team Leader" : `Member ${i}`,
        name: memberName,
        usn: memberUsn,
        phone: memberPhone,
      });
    }

    if (!req.file) {
      return res
        .status(400)
        .json({ message: "Payment screenshot is required" });
    }

    const records = readRegistrations();

    const registration = {
      id: Date.now(),
      createdAt: new Date().toISOString(),
      email,
      teamName,
      teamPhone,
      teamCollege,
      domain,
      teamSize: size,
      transactionId,
      paymentScreenshot: `/uploads/${req.file.filename}`,
      members,
    };

    records.push(registration);
    fs.writeFileSync(dbFile, JSON.stringify(records, null, 2), "utf-8");

    return res
      .status(201)
      .json({ message: "Registration saved", data: registration });
  } catch (error) {
    return res
      .status(500)
      .json({ message: "Server error", error: error.message });
  }
});

app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError || err) {
    return res.status(400).json({ message: err.message || "Upload error" });
  }
  return next();
});

function startServer(port) {
  const server = app.listen(port, () => {
    console.log(`Server running on http://localhost:${port}`);
  });

  server.on("error", (error) => {
    if (error.code === "EADDRINUSE") {
      const nextPort = port + 1;
      console.log(
        `Port ${port} is busy. Trying http://localhost:${nextPort}...`,
      );
      startServer(nextPort);
      return;
    }
    throw error;
  });
}

startServer(PORT);
