// ============================================================
//  MiciMart - server.js  (PHẦN 1/4)
//  Paste 4 phần nối tiếp nhau vào 1 file
// ============================================================
require("dotenv").config();
const express = require("express");
const { Pool } = require("pg");
const bcrypt = require("bcryptjs");
const jwt = require("jsonwebtoken");
const cors = require("cors");
const multer = require("multer");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

const JWT_SECRET = process.env.JWT_SECRET || "micimart2026secret";
const JWT_EXPIRES = process.env.JWT_EXPIRES || "8h";
const PORT = process.env.PORT || 3000;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
});

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl:
    process.env.NODE_ENV === "production"
      ? { rejectUnauthorized: false }
      : false,
});

// ── INIT DB ───────────────────────────────────────────────
async function initDB() {
  const c = await pool.connect();
  try {
    await c.query(`
      CREATE TABLE IF NOT EXISTS roles (
        id           SERIAL PRIMARY KEY,
        name         VARCHAR(50)  UNIQUE NOT NULL,
        display_name VARCHAR(100) NOT NULL,
        permissions  TEXT NOT NULL DEFAULT '[]'
      );
      CREATE TABLE IF NOT EXISTS users (
        id            SERIAL PRIMARY KEY,
        username      VARCHAR(50)  UNIQUE NOT NULL,
        password_hash VARCHAR(255) NOT NULL,
        full_name     VARCHAR(100) NOT NULL,
        email         VARCHAR(100),
        phone         VARCHAR(20),
        role_id       INT NOT NULL DEFAULT 7,
        is_active     BOOLEAN DEFAULT TRUE,
        avatar        VARCHAR(10) DEFAULT '👤',
        last_login    TIMESTAMP,
        created_at    TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS products (
        id            SERIAL PRIMARY KEY,
        code          VARCHAR(30)  UNIQUE NOT NULL,
        name          VARCHAR(255) NOT NULL,
        category      VARCHAR(100) NOT NULL DEFAULT 'Khác',
        subcategory   VARCHAR(100),
        unit          VARCHAR(50)  NOT NULL DEFAULT 'Cái',
        cost_price    BIGINT NOT NULL DEFAULT 0,
        selling_price BIGINT NOT NULL DEFAULT 0,
        discount_percent INT DEFAULT 0,
        stock         INT    NOT NULL DEFAULT 0,
        min_stock     INT    NOT NULL DEFAULT 10,
        icon          VARCHAR(10) DEFAULT '📦',
        is_active     BOOLEAN DEFAULT TRUE,
        created_at    TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS orders (
        id              SERIAL PRIMARY KEY,
        order_code      VARCHAR(30) UNIQUE,
        customer_name   VARCHAR(100) DEFAULT 'Khách lẻ',
        customer_id     INT,
        subtotal        BIGINT NOT NULL DEFAULT 0,
        discount        BIGINT NOT NULL DEFAULT 0,
        vat             BIGINT NOT NULL DEFAULT 0,
        total           BIGINT NOT NULL DEFAULT 0,
        payment_method  VARCHAR(20) DEFAULT 'cash',
        cashier_id      INT,
        status          VARCHAR(20) DEFAULT 'done',
        delivery_status VARCHAR(20) DEFAULT 'none',
        delivery_id     INT,
        delivery_address TEXT,
        notes           TEXT,
        voucher_code    VARCHAR(50),
        created_at      TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS order_items (
        id           SERIAL PRIMARY KEY,
        order_id     INT NOT NULL REFERENCES orders(id) ON DELETE CASCADE,
        product_id   INT REFERENCES products(id) ON DELETE SET NULL,
        product_name VARCHAR(255) NOT NULL,
        qty          INT   NOT NULL DEFAULT 1,
        price        BIGINT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS customer_profiles (
        id          SERIAL PRIMARY KEY,
        user_id     INT UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        rank        VARCHAR(20) DEFAULT 'Thường',
        points      INT DEFAULT 0,
        total_spent BIGINT DEFAULT 0
      );
      CREATE TABLE IF NOT EXISTS vouchers (
        id           SERIAL PRIMARY KEY,
        code         VARCHAR(50) UNIQUE NOT NULL,
        description  VARCHAR(255),
        discount_type VARCHAR(20) DEFAULT 'percent',
        discount_value INT NOT NULL DEFAULT 0,
        min_order    BIGINT DEFAULT 0,
        max_uses     INT DEFAULT 100,
        used_count   INT DEFAULT 0,
        is_active    BOOLEAN DEFAULT TRUE,
        expires_at   TIMESTAMP,
        created_at   TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS user_vouchers (
        id         SERIAL PRIMARY KEY,
        user_id    INT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        voucher_id INT NOT NULL REFERENCES vouchers(id) ON DELETE CASCADE,
        used       BOOLEAN DEFAULT FALSE,
        assigned_at TIMESTAMP DEFAULT NOW(),
        UNIQUE(user_id, voucher_id)
      );
      CREATE TABLE IF NOT EXISTS attendance (
        id         SERIAL PRIMARY KEY,
        user_id    INT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        date       DATE NOT NULL DEFAULT CURRENT_DATE,
        checkin    TIMESTAMP,
        checkout   TIMESTAMP,
        status     VARCHAR(20) DEFAULT 'present',
        UNIQUE(user_id, date)
      );
      CREATE TABLE IF NOT EXISTS salary (
        id           SERIAL PRIMARY KEY,
        user_id      INT UNIQUE NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        base_salary  BIGINT DEFAULT 0,
        bonus        BIGINT DEFAULT 0,
        deduction    BIGINT DEFAULT 0,
        updated_at   TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS feedback (
        id         SERIAL PRIMARY KEY,
        user_id    INT REFERENCES users(id) ON DELETE SET NULL,
        name       VARCHAR(100),
        content    TEXT NOT NULL,
        rating     INT DEFAULT 5,
        status     VARCHAR(20) DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT NOW()
      );
      CREATE TABLE IF NOT EXISTS delivery_addresses (
        id         SERIAL PRIMARY KEY,
        user_id    INT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
        address    TEXT NOT NULL,
        label      VARCHAR(50),
        is_default BOOLEAN DEFAULT FALSE,
        created_at TIMESTAMP DEFAULT NOW()
      );
    `);

    // Seed roles
    await c.query(`
      INSERT INTO roles (id, name, display_name, permissions) VALUES
        (1,'admin',       'Quản trị viên',      '["dashboard","pos","products","orders","inventory","customers","vouchers","reports","staff","user_management","delivery_mgmt","attendance","salary","feedback"]'),
        (2,'manager',     'Quản lý cửa hàng',   '["dashboard","pos","products","orders","inventory","customers","vouchers","reports","staff","delivery_mgmt","attendance","salary","feedback"]'),
        (3,'cashier',     'Thu ngân',            '["pos","orders","customers","my_shift","my_orders","my_vouchers","my_profile","attendance"]'),
        (4,'warehouse',   'Nhân viên kho',       '["inventory","products","my_profile","attendance"]'),
        (5,'delivery',    'Nhân viên giao hàng', '["delivery","my_profile","attendance"]'),
        (6,'salesperson', 'Nhân viên bán hàng',  '["pos","products","inventory","my_profile","attendance"]'),
        (7,'customer',    'Khách hàng',          '["my_orders","my_vouchers","my_profile"]')
      ON CONFLICT (name) DO NOTHING;
    `);
    await c.query(
      `SELECT setval('roles_id_seq', GREATEST((SELECT MAX(id) FROM roles), 1));`,
    );

    // Seed admin
    const ck = await c.query(`SELECT id FROM users WHERE username='admin12'`);
    if (!ck.rows.length) {
      const h = await bcrypt.hash("admin12", 10);
      await c.query(
        `INSERT INTO users (username,password_hash,full_name,email,role_id)
         VALUES ('admin12',$1,'Admin MiciMart','admin@micimart.vn',1)`,
        [h],
      );
      console.log("🔑 Tài khoản admin: admin12 / admin12");
    }

    // Seed vouchers
    await c.query(`
      INSERT INTO vouchers (code,description,discount_type,discount_value,min_order,max_uses) VALUES
        ('WELCOME10','Giảm 10% đơn đầu tiên','percent',10,0,1000),
        ('MICIMART20','Giảm 20% đơn từ 200k','percent',20,200000,500),
        ('FREESHIP','Miễn phí giao hàng','fixed',30000,100000,200)
      ON CONFLICT (code) DO NOTHING;
    `);

    // Seed products
    await c.query(`
      INSERT INTO products (code,name,category,unit,cost_price,selling_price,stock,min_stock,icon) VALUES
        ('SP001','Sữa tươi TH True Milk 1L','Sữa & chế phẩm từ sữa','Hộp',28000,35000,124,20,'🥛'),
        ('SP002','Gạo ST25 5kg','Thực phẩm tươi sống','Kg',65000,85000,48,15,'🌾'),
        ('SP003','Nước ngọt Pepsi 1.5L','Đồ uống','Chai',18000,25000,200,30,'🥤'),
        ('SP004','Bánh Oreo 133g','Bánh kẹo & snack','Gói',22000,30000,87,25,'🍪'),
        ('SP005','Trứng gà ta (vỉ 10)','Thực phẩm tươi sống','Vỉ',32000,42000,60,20,'🥚'),
        ('SP006','Bia Tiger thùng 24','Đồ uống','Thùng',280000,360000,30,10,'🍺'),
        ('SP007','Nước mắm Phú Quốc 500ml','Gia vị','Chai',35000,48000,75,15,'🫙'),
        ('SP008','Dầu ăn Neptune 2L','Gia vị','Chai',65000,82000,40,15,'🫒'),
        ('SP009','Xà phòng Dove 100g','Chăm sóc cá nhân','Cái',18000,26000,8,20,'🧼'),
        ('SP010','Nước rửa bát Sunlight 750ml','Hóa phẩm','Chai',22000,32000,55,15,'🧴'),
        ('SP011','Mì Hảo Hảo thùng 30 gói','Thực phẩm tươi sống','Thùng',95000,125000,22,10,'🍜'),
        ('SP012','Sữa chua Vinamilk lốc 4','Sữa & chế phẩm từ sữa','Lốc',28000,36000,6,15,'🍦')
      ON CONFLICT (code) DO NOTHING;
    `);

    console.log("✅ Database sẵn sàng");
  } finally {
    c.release();
  }
}

// ── MIDDLEWARE ────────────────────────────────────────────
const auth = (req, res, next) => {
  const h = req.headers.authorization;
  if (!h?.startsWith("Bearer "))
    return res.status(401).json({ error: "Chưa đăng nhập" });
  try {
    req.user = jwt.verify(h.split(" ")[1], JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: "Phiên đăng nhập hết hạn" });
  }
};
const can =
  (...roles) =>
  (req, res, next) => {
    if (!roles.includes(req.user.role))
      return res
        .status(403)
        .json({ error: "Không có quyền thực hiện thao tác này" });
    next();
  };

// ── AUTH ──────────────────────────────────────────────────
app.post("/api/auth/register", async (req, res) => {
  const { username, password, full_name, email, phone } = req.body;
  if (!username || !password || !full_name)
    return res.status(400).json({ error: "Vui lòng điền đầy đủ thông tin" });
  if (password.length < 6)
    return res.status(400).json({ error: "Mật khẩu phải ít nhất 6 ký tự" });
  try {
    const hash = await bcrypt.hash(password, 10);
    const r = await pool.query(
      `INSERT INTO users (username,password_hash,full_name,email,phone,role_id)
       VALUES ($1,$2,$3,$4,$5,7) RETURNING id`,
      [username.trim(), hash, full_name.trim(), email || null, phone || null],
    );
    await pool.query("INSERT INTO customer_profiles (user_id) VALUES ($1)", [
      r.rows[0].id,
    ]);
    res.status(201).json({ message: "Đăng ký thành công! Hãy đăng nhập." });
  } catch (e) {
    if (e.code === "23505")
      return res.status(409).json({ error: "Tên đăng nhập đã tồn tại" });
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.post("/api/auth/login", async (req, res) => {
  const { username, password } = req.body;
  if (!username || !password)
    return res.status(400).json({ error: "Vui lòng nhập thông tin đăng nhập" });
  try {
    const r = await pool.query(
      `SELECT u.*, ro.name AS role, ro.display_name AS role_display, ro.permissions
       FROM users u JOIN roles ro ON u.role_id=ro.id WHERE u.username=$1`,
      [username.trim()],
    );
    if (!r.rows.length)
      return res.status(401).json({ error: "Sai tên đăng nhập hoặc mật khẩu" });
    const u = r.rows[0];
    if (!u.is_active)
      return res
        .status(403)
        .json({ error: "Tài khoản đã bị khóa. Liên hệ Admin." });
    if (!(await bcrypt.compare(password, u.password_hash)))
      return res.status(401).json({ error: "Sai tên đăng nhập hoặc mật khẩu" });
    await pool.query("UPDATE users SET last_login=NOW() WHERE id=$1", [u.id]);
    const perms =
      typeof u.permissions === "string"
        ? JSON.parse(u.permissions)
        : u.permissions;
    const payload = {
      id: u.id,
      username: u.username,
      full_name: u.full_name,
      role: u.role,
      role_display: u.role_display,
      permissions: perms,
    };
    res.json({
      token: jwt.sign(payload, JWT_SECRET, { expiresIn: JWT_EXPIRES }),
      user: payload,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.get("/api/auth/me", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT u.id,u.username,u.full_name,u.email,u.phone,u.last_login,
            ro.name AS role, ro.display_name AS role_display, ro.permissions
     FROM users u JOIN roles ro ON u.role_id=ro.id WHERE u.id=$1`,
    [req.user.id],
  );
  if (!r.rows.length) return res.status(404).json({ error: "Không tìm thấy" });
  const u = r.rows[0];
  u.permissions =
    typeof u.permissions === "string"
      ? JSON.parse(u.permissions)
      : u.permissions;
  res.json(u);
});

// ── KẾT THÚC PHẦN 1/4 ──
// → Tiếp tục paste PHẦN 2/4 ngay bên dưới
// ── PHẦN 2/4: Users + Products + Orders ──────────────────

// ── USERS ─────────────────────────────────────────────────
app.get("/api/users", auth, can("admin"), async (req, res) => {
  const r = await pool.query(
    `SELECT u.id,u.username,u.full_name,u.email,u.phone,u.is_active,u.last_login,u.created_at,
            ro.id AS role_id, ro.name AS role, ro.display_name AS role_display
     FROM users u JOIN roles ro ON u.role_id=ro.id ORDER BY u.created_at DESC`,
  );
  res.json(r.rows);
});

app.get("/api/roles", auth, can("admin", "manager"), async (req, res) => {
  const r = await pool.query(
    "SELECT id,name,display_name FROM roles ORDER BY id",
  );
  res.json(r.rows);
});

app.post("/api/users", auth, can("admin", "manager"), async (req, res) => {
  const { username, password, full_name, email, phone, role_id } = req.body;
  if (!username || !password || !full_name || !role_id)
    return res.status(400).json({ error: "Thiếu thông tin bắt buộc" });
  try {
    const hash = await bcrypt.hash(password, 10);
    await pool.query(
      `INSERT INTO users (username,password_hash,full_name,email,phone,role_id)
       VALUES ($1,$2,$3,$4,$5,$6)`,
      [
        username.trim(),
        hash,
        full_name.trim(),
        email || null,
        phone || null,
        role_id,
      ],
    );
    res.status(201).json({ message: "Tạo tài khoản thành công" });
  } catch (e) {
    if (e.code === "23505")
      return res.status(409).json({ error: "Tên đăng nhập đã tồn tại" });
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.put("/api/users/:id/role", auth, can("admin"), async (req, res) => {
  const uid = parseInt(req.params.id);
  if (uid === req.user.id)
    return res.status(400).json({ error: "Không thể tự đổi vai trò của mình" });
  await pool.query("UPDATE users SET role_id=$1 WHERE id=$2", [
    req.body.role_id,
    uid,
  ]);
  res.json({ message: "Cập nhật vai trò thành công" });
});

app.put("/api/users/:id/toggle", auth, can("admin"), async (req, res) => {
  const uid = parseInt(req.params.id);
  if (uid === req.user.id)
    return res
      .status(400)
      .json({ error: "Không thể khóa tài khoản đang đăng nhập" });
  const r = await pool.query(
    "UPDATE users SET is_active=NOT is_active WHERE id=$1 RETURNING is_active",
    [uid],
  );
  res.json({
    message: r.rows[0].is_active ? "Đã mở khóa tài khoản" : "Đã khóa tài khoản",
  });
});

app.delete("/api/users/:id", auth, can("admin"), async (req, res) => {
  const uid = parseInt(req.params.id);
  if (uid === req.user.id)
    return res
      .status(400)
      .json({ error: "Không thể xóa tài khoản đang đăng nhập" });
  await pool.query("DELETE FROM users WHERE id=$1", [uid]);
  res.json({ message: "Đã xóa tài khoản" });
});

// Danh sách nhân viên giao hàng
app.get(
  "/api/users/delivery-staff",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const r = await pool.query(
      `SELECT u.id,u.full_name,u.phone FROM users u
     JOIN roles ro ON u.role_id=ro.id
     WHERE ro.name='delivery' AND u.is_active=true ORDER BY u.full_name`,
    );
    res.json(r.rows);
  },
);

// ── PRODUCTS ──────────────────────────────────────────────
app.get("/api/products/public", async (req, res) => {
  const r = await pool.query(
    "SELECT * FROM products WHERE is_active=true ORDER BY id",
  );
  res.json(r.rows);
});

app.get("/api/products", auth, async (req, res) => {
  const r = await pool.query(
    "SELECT * FROM products WHERE is_active=true ORDER BY id",
  );
  res.json(r.rows);
});

app.post(
  "/api/products",
  auth,
  can("admin", "manager", "salesperson"),
  async (req, res) => {
    const {
      code,
      name,
      category,
      subcategory,
      unit,
      cost_price,
      selling_price,
      discount_percent,
      stock,
      min_stock,
      icon,
    } = req.body;
    try {
      await pool.query(
        `INSERT INTO products (code,name,category,subcategory,unit,cost_price,selling_price,discount_percent,stock,min_stock,icon)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
        [
          code,
          name,
          category || "Khác",
          subcategory || null,
          unit || "Cái",
          cost_price || 0,
          selling_price || 0,
          discount_percent || 0,
          stock || 0,
          min_stock || 10,
          icon || "📦",
        ],
      );
      res.status(201).json({ message: "Thêm sản phẩm thành công" });
    } catch (e) {
      if (e.code === "23505")
        return res.status(409).json({ error: "Mã sản phẩm đã tồn tại" });
      res.status(500).json({ error: "Lỗi hệ thống" });
    }
  },
);

app.put(
  "/api/products/:id",
  auth,
  can("admin", "manager", "salesperson"),
  async (req, res) => {
    const {
      name,
      category,
      subcategory,
      unit,
      cost_price,
      selling_price,
      discount_percent,
      stock,
      min_stock,
      icon,
    } = req.body;
    await pool.query(
      `UPDATE products SET name=$1,category=$2,subcategory=$3,unit=$4,cost_price=$5,
     selling_price=$6,discount_percent=$7,stock=$8,min_stock=$9,icon=$10 WHERE id=$11`,
      [
        name,
        category,
        subcategory || null,
        unit,
        cost_price,
        selling_price,
        discount_percent || 0,
        stock,
        min_stock,
        icon,
        req.params.id,
      ],
    );
    res.json({ message: "Cập nhật thành công" });
  },
);

app.delete(
  "/api/products/:id",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    await pool.query("UPDATE products SET is_active=false WHERE id=$1", [
      req.params.id,
    ]);
    res.json({ message: "Đã xóa sản phẩm" });
  },
);

// Import Excel sản phẩm
app.post(
  "/api/import/products",
  auth,
  can("admin", "manager"),
  upload.single("file"),
  async (req, res) => {
    if (!req.file)
      return res.status(400).json({ error: "Vui lòng chọn file Excel" });
    try {
      const wb = XLSX.read(req.file.buffer, { type: "buffer" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      let ok = 0,
        skip = 0;
      for (const row of rows) {
        const code = String(row["Mã SP"] || row["code"] || "").trim();
        const name = String(row["Tên sản phẩm"] || row["name"] || "").trim();
        if (!code || !name) {
          skip++;
          continue;
        }
        try {
          await pool.query(
            `INSERT INTO products (code,name,category,unit,cost_price,selling_price,stock,min_stock,icon)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9)
           ON CONFLICT (code) DO UPDATE SET
             name=$2,category=$3,unit=$4,cost_price=$5,selling_price=$6,stock=$7,min_stock=$8,icon=$9`,
            [
              code,
              name,
              String(row["Danh mục"] || row["category"] || "Khác").trim(),
              String(row["Đơn vị"] || row["unit"] || "Cái").trim(),
              parseInt(row["Giá nhập"] || row["cost_price"] || 0),
              parseInt(row["Giá bán"] || row["selling_price"] || 0),
              parseInt(row["Tồn kho"] || row["stock"] || 0),
              parseInt(row["Tồn tối thiểu"] || row["min_stock"] || 10),
              String(row["Icon"] || row["icon"] || "📦").trim(),
            ],
          );
          ok++;
        } catch (e) {
          skip++;
        }
      }
      res.json({
        message: `Import thành công ${ok} sản phẩm${skip ? `, bỏ qua ${skip} dòng` : ""}`,
        ok,
        skip,
      });
    } catch (e) {
      res.status(400).json({ error: "Lỗi đọc file Excel" });
    }
  },
);

// Import Excel nhân viên
app.post(
  "/api/import/staff",
  auth,
  can("admin"),
  upload.single("file"),
  async (req, res) => {
    if (!req.file)
      return res.status(400).json({ error: "Vui lòng chọn file Excel" });
    try {
      const wb = XLSX.read(req.file.buffer, { type: "buffer" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      const roleMap = {
        "Quản trị viên": 1,
        "Quản lý cửa hàng": 2,
        "Thu ngân": 3,
        "Nhân viên kho": 4,
        "Nhân viên giao hàng": 5,
        "Nhân viên bán hàng": 6,
        "Khách hàng": 7,
      };
      let ok = 0,
        skip = 0;
      for (const row of rows) {
        const username = String(
          row["Tên đăng nhập"] || row["username"] || "",
        ).trim();
        const password = String(
          row["Mật khẩu"] || row["password"] || "",
        ).trim();
        const full_name = String(
          row["Họ và tên"] || row["full_name"] || "",
        ).trim();
        if (!username || !password || !full_name) {
          skip++;
          continue;
        }
        try {
          const hash = await bcrypt.hash(password, 10);
          await pool.query(
            `INSERT INTO users (username,password_hash,full_name,email,phone,role_id)
           VALUES ($1,$2,$3,$4,$5,$6)
           ON CONFLICT (username) DO UPDATE SET
             password_hash=$2,full_name=$3,email=$4,phone=$5,role_id=$6`,
            [
              username,
              hash,
              full_name,
              String(row["Email"] || row["email"] || "").trim() || null,
              String(row["Số điện thoại"] || row["phone"] || "").trim() || null,
              roleMap[
                String(row["Vai trò"] || row["role"] || "Thu ngân").trim()
              ] || 3,
            ],
          );
          ok++;
        } catch (e) {
          skip++;
        }
      }
      res.json({
        message: `Import thành công ${ok} nhân viên${skip ? `, bỏ qua ${skip} dòng` : ""}`,
        ok,
        skip,
      });
    } catch (e) {
      res.status(400).json({ error: "Lỗi đọc file Excel" });
    }
  },
);

// ── ORDERS ────────────────────────────────────────────────
app.get("/api/orders", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT o.*, u.full_name AS cashier_name
     FROM orders o LEFT JOIN users u ON o.cashier_id=u.id
     ORDER BY o.created_at DESC LIMIT 300`,
  );
  res.json(r.rows);
});

app.post(
  "/api/orders",
  auth,
  can("admin", "manager", "cashier", "salesperson"),
  async (req, res) => {
    const {
      customer_name,
      customer_id,
      items,
      discount = 0,
      payment_method = "cash",
      delivery_address,
      notes,
      voucher_code,
    } = req.body;
    if (!items?.length)
      return res.status(400).json({ error: "Giỏ hàng trống" });
    const c = await pool.connect();
    try {
      await c.query("BEGIN");
      const subtotal = items.reduce((s, i) => s + i.price * i.qty, 0);
      const vat = Math.round((subtotal - discount) * 0.1);
      const total = subtotal - discount + vat;
      const code = "DH" + Date.now();
      const delivStatus = delivery_address ? "pending" : "none";
      const r = await c.query(
        `INSERT INTO orders (order_code,customer_name,customer_id,subtotal,discount,vat,total,
        payment_method,cashier_id,delivery_status,delivery_address,notes,voucher_code)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13) RETURNING id`,
        [
          code,
          customer_name || "Khách lẻ",
          customer_id || null,
          subtotal,
          discount,
          vat,
          total,
          payment_method,
          req.user.id,
          delivStatus,
          delivery_address || null,
          notes || null,
          voucher_code || null,
        ],
      );
      const oid = r.rows[0].id;
      for (const item of items) {
        await c.query(
          `INSERT INTO order_items (order_id,product_id,product_name,qty,price)
         VALUES ($1,$2,$3,$4,$5)`,
          [oid, item.product_id || null, item.name, item.qty, item.price],
        );
        if (item.product_id)
          await c.query("UPDATE products SET stock=stock-$1 WHERE id=$2", [
            item.qty,
            item.product_id,
          ]);
      }
      // Cập nhật voucher nếu dùng
      if (voucher_code) {
        await c.query(
          `UPDATE vouchers SET used_count=used_count+1 WHERE code=$1`,
          [voucher_code],
        );
        if (customer_id) {
          await c.query(
            `UPDATE user_vouchers SET used=true
           WHERE user_id=$1 AND voucher_id=(SELECT id FROM vouchers WHERE code=$2)`,
            [customer_id, voucher_code],
          );
        }
      }
      // Cập nhật điểm & tổng chi khách hàng
      if (customer_id) {
        const pts = Math.floor(total / 10000);
        await c.query(
          `UPDATE customer_profiles SET total_spent=total_spent+$1, points=points+$2 WHERE user_id=$3`,
          [total, pts, customer_id],
        );
      }
      await c.query("COMMIT");
      res
        .status(201)
        .json({ message: "Thanh toán thành công", order_id: oid, total });
    } catch (e) {
      await c.query("ROLLBACK");
      console.error(e);
      res.status(500).json({ error: "Lỗi xử lý đơn hàng" });
    } finally {
      c.release();
    }
  },
);

app.put(
  "/api/orders/:id/status",
  auth,
  can("admin", "manager", "cashier"),
  async (req, res) => {
    await pool.query("UPDATE orders SET status=$1 WHERE id=$2", [
      req.body.status,
      req.params.id,
    ]);
    res.json({ message: "Cập nhật trạng thái thành công" });
  },
);

// Đơn hàng theo ca (cashier)
app.get("/api/orders/my-shift", auth, async (req, res) => {
  const today = new Date().toISOString().split("T")[0];
  const r = await pool.query(
    `SELECT o.*, u.full_name AS cashier_name FROM orders o
     LEFT JOIN users u ON o.cashier_id=u.id
     WHERE o.cashier_id=$1 AND DATE(o.created_at)=$2
     ORDER BY o.created_at DESC`,
    [req.user.id, today],
  );
  res.json(r.rows);
});

// Đơn hàng của khách hàng
app.get("/api/orders/my-orders", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT o.* FROM orders o WHERE o.customer_id=$1 ORDER BY o.created_at DESC`,
    [req.user.id],
  );
  res.json(r.rows);
});

// Đơn chờ giao
app.get(
  "/api/orders/delivery-pending",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const r = await pool.query(
      `SELECT o.*, u.full_name AS cashier_name FROM orders o
     LEFT JOIN users u ON o.cashier_id=u.id
     WHERE o.delivery_status='pending' ORDER BY o.created_at ASC`,
    );
    res.json(r.rows);
  },
);

// Đơn của shipper
app.get("/api/orders/my-deliveries", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT o.* FROM orders o WHERE o.delivery_id=$1
     AND o.delivery_status IN ('shipping','pending')
     ORDER BY o.created_at ASC`,
    [req.user.id],
  );
  res.json(r.rows);
});

// Phân công shipper
app.put(
  "/api/orders/:id/assign-delivery",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    await pool.query(
      `UPDATE orders SET delivery_id=$1, delivery_status='shipping' WHERE id=$2`,
      [req.body.delivery_id, req.params.id],
    );
    res.json({ message: "Đã phân công giao hàng" });
  },
);

// Cập nhật trạng thái giao hàng
app.put("/api/orders/:id/delivery", auth, async (req, res) => {
  const { delivery_status } = req.body;
  await pool.query(`UPDATE orders SET delivery_status=$1 WHERE id=$2`, [
    delivery_status,
    req.params.id,
  ]);
  res.json({ message: "Cập nhật giao hàng thành công" });
});

// Địa chỉ giao hàng
app.get("/api/addresses", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT * FROM delivery_addresses WHERE user_id=$1 ORDER BY is_default DESC, id DESC`,
    [req.user.id],
  );
  res.json(r.rows);
});

app.post("/api/addresses", auth, async (req, res) => {
  const { address, label, is_default } = req.body;
  if (is_default) {
    await pool.query(
      `UPDATE delivery_addresses SET is_default=false WHERE user_id=$1`,
      [req.user.id],
    );
  }
  await pool.query(
    `INSERT INTO delivery_addresses (user_id,address,label,is_default) VALUES ($1,$2,$3,$4)`,
    [req.user.id, address, label || null, is_default || false],
  );
  res.status(201).json({ message: "Đã thêm địa chỉ" });
});

app.delete("/api/addresses/:id", auth, async (req, res) => {
  await pool.query(
    `DELETE FROM delivery_addresses WHERE id=$1 AND user_id=$2`,
    [req.params.id, req.user.id],
  );
  res.json({ message: "Đã xóa địa chỉ" });
});

// ── KẾT THÚC PHẦN 2/4 ──
// → Tiếp tục paste PHẦN 3/4 ngay bên dưới
// ── PHẦN 3/4: Vouchers + Attendance + Salary + Feedback ──

// ── VOUCHERS ──────────────────────────────────────────────
app.get("/api/vouchers", auth, can("admin", "manager"), async (req, res) => {
  const r = await pool.query(
    `SELECT v.*, COUNT(uv.id) AS assigned_count
     FROM vouchers v
     LEFT JOIN user_vouchers uv ON v.id=uv.voucher_id
     GROUP BY v.id ORDER BY v.created_at DESC`,
  );
  res.json(r.rows);
});

app.post("/api/vouchers", auth, can("admin", "manager"), async (req, res) => {
  const {
    code,
    description,
    discount_type,
    discount_value,
    min_order,
    max_uses,
    expires_at,
  } = req.body;
  if (!code || !discount_value)
    return res.status(400).json({ error: "Thiếu thông tin bắt buộc" });
  try {
    await pool.query(
      `INSERT INTO vouchers (code,description,discount_type,discount_value,min_order,max_uses,expires_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7)`,
      [
        code.toUpperCase().trim(),
        description || null,
        discount_type || "percent",
        discount_value,
        min_order || 0,
        max_uses || 100,
        expires_at || null,
      ],
    );
    res.status(201).json({ message: "Tạo voucher thành công" });
  } catch (e) {
    if (e.code === "23505")
      return res.status(409).json({ error: "Mã voucher đã tồn tại" });
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.put(
  "/api/vouchers/:id",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const {
      description,
      discount_value,
      min_order,
      max_uses,
      is_active,
      expires_at,
    } = req.body;
    await pool.query(
      `UPDATE vouchers SET description=$1,discount_value=$2,min_order=$3,
     max_uses=$4,is_active=$5,expires_at=$6 WHERE id=$7`,
      [
        description,
        discount_value,
        min_order,
        max_uses,
        is_active,
        expires_at || null,
        req.params.id,
      ],
    );
    res.json({ message: "Cập nhật voucher thành công" });
  },
);

// Kiểm tra voucher hợp lệ
app.post("/api/vouchers/check", auth, async (req, res) => {
  const { code, order_total } = req.body;
  if (!code) return res.status(400).json({ error: "Thiếu mã voucher" });
  try {
    const r = await pool.query(
      `SELECT * FROM vouchers WHERE code=$1 AND is_active=true`,
      [code.toUpperCase().trim()],
    );
    if (!r.rows.length)
      return res
        .status(404)
        .json({ error: "Mã voucher không tồn tại hoặc đã hết hiệu lực" });
    const v = r.rows[0];
    if (v.used_count >= v.max_uses)
      return res.status(400).json({ error: "Voucher đã hết lượt sử dụng" });
    if (v.expires_at && new Date(v.expires_at) < new Date())
      return res.status(400).json({ error: "Voucher đã hết hạn" });
    if (order_total < v.min_order)
      return res.status(400).json({
        error: `Đơn hàng tối thiểu ${Number(v.min_order).toLocaleString("vi-VN")}đ`,
      });
    const discount =
      v.discount_type === "percent"
        ? Math.round((order_total * v.discount_value) / 100)
        : v.discount_value;
    res.json({ valid: true, voucher: v, discount });
  } catch (e) {
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

// Gán voucher cho user
app.post(
  "/api/vouchers/assign",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { user_id, voucher_id } = req.body;
    try {
      await pool.query(
        `INSERT INTO user_vouchers (user_id,voucher_id) VALUES ($1,$2) ON CONFLICT DO NOTHING`,
        [user_id, voucher_id],
      );
      res.json({ message: "Đã gán voucher thành công" });
    } catch (e) {
      res.status(500).json({ error: "Lỗi hệ thống" });
    }
  },
);

// Gán voucher cho nhiều user
app.post(
  "/api/vouchers/assign-bulk",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { voucher_id, role_id } = req.body;
    try {
      let users;
      if (role_id) {
        users = await pool.query(
          `SELECT id FROM users WHERE role_id=$1 AND is_active=true`,
          [role_id],
        );
      } else {
        users = await pool.query(`SELECT id FROM users WHERE is_active=true`);
      }
      let ok = 0;
      for (const u of users.rows) {
        try {
          await pool.query(
            `INSERT INTO user_vouchers (user_id,voucher_id) VALUES ($1,$2) ON CONFLICT DO NOTHING`,
            [u.id, voucher_id],
          );
          ok++;
        } catch (e) {
          /* bỏ qua */
        }
      }
      res.json({ message: `Đã gán voucher cho ${ok} người dùng` });
    } catch (e) {
      res.status(500).json({ error: "Lỗi hệ thống" });
    }
  },
);

// Voucher của user
app.get("/api/my-vouchers", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT v.*, uv.used, uv.assigned_at FROM user_vouchers uv
     JOIN vouchers v ON uv.voucher_id=v.id
     WHERE uv.user_id=$1 ORDER BY uv.assigned_at DESC`,
    [req.user.id],
  );
  res.json(r.rows);
});

// ── ATTENDANCE ─────────────────────────────────────────────
app.post("/api/attendance/checkin", auth, async (req, res) => {
  const today = new Date().toISOString().split("T")[0];
  try {
    await pool.query(
      `INSERT INTO attendance (user_id,date,checkin,status)
       VALUES ($1,$2,NOW(),'present')
       ON CONFLICT (user_id,date) DO UPDATE SET checkin=NOW()`,
      [req.user.id, today],
    );
    res.json({
      message: "Điểm danh vào ca thành công",
      time: new Date().toLocaleTimeString("vi-VN"),
    });
  } catch (e) {
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.post("/api/attendance/checkout", auth, async (req, res) => {
  const today = new Date().toISOString().split("T")[0];
  try {
    const r = await pool.query(
      `UPDATE attendance SET checkout=NOW() WHERE user_id=$1 AND date=$2 RETURNING *`,
      [req.user.id, today],
    );
    if (!r.rows.length)
      return res.status(400).json({ error: "Chưa điểm danh vào ca" });
    res.json({
      message: "Kết thúc ca thành công",
      time: new Date().toLocaleTimeString("vi-VN"),
    });
  } catch (e) {
    res.status(500).json({ error: "Lỗi hệ thống" });
  }
});

app.get("/api/attendance", auth, can("admin", "manager"), async (req, res) => {
  const { date } = req.query;
  const target = date || new Date().toISOString().split("T")[0];
  const r = await pool.query(
    `SELECT a.*, u.full_name, u.username, ro.display_name AS role_display
     FROM attendance a
     JOIN users u ON a.user_id=u.id
     JOIN roles ro ON u.role_id=ro.id
     WHERE a.date=$1 ORDER BY a.checkin ASC`,
    [target],
  );
  res.json(r.rows);
});

app.get("/api/attendance/my", auth, async (req, res) => {
  const r = await pool.query(
    `SELECT * FROM attendance WHERE user_id=$1
     ORDER BY date DESC LIMIT 30`,
    [req.user.id],
  );
  res.json(r.rows);
});

// ── SALARY ────────────────────────────────────────────────
app.get("/api/salary", auth, can("admin", "manager"), async (req, res) => {
  const r = await pool.query(
    `SELECT s.*, u.full_name, u.username, ro.display_name AS role_display
     FROM salary s
     JOIN users u ON s.user_id=u.id
     JOIN roles ro ON u.role_id=ro.id
     ORDER BY u.full_name`,
  );
  res.json(r.rows);
});

app.get("/api/salary/:userId", auth, async (req, res) => {
  const uid = parseInt(req.params.userId);
  if (req.user.id !== uid && !["admin", "manager"].includes(req.user.role))
    return res.status(403).json({ error: "Không có quyền" });
  const r = await pool.query(
    `SELECT s.*, u.full_name FROM salary s
     JOIN users u ON s.user_id=u.id WHERE s.user_id=$1`,
    [uid],
  );
  if (!r.rows.length)
    return res.json({ user_id: uid, base_salary: 0, bonus: 0, deduction: 0 });
  res.json(r.rows[0]);
});

app.put(
  "/api/salary/:userId",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { base_salary, bonus, deduction } = req.body;
    await pool.query(
      `INSERT INTO salary (user_id,base_salary,bonus,deduction,updated_at)
     VALUES ($1,$2,$3,$4,NOW())
     ON CONFLICT (user_id) DO UPDATE SET
       base_salary=$2,bonus=$3,deduction=$4,updated_at=NOW()`,
      [req.params.userId, base_salary || 0, bonus || 0, deduction || 0],
    );
    res.json({ message: "Cập nhật lương thành công" });
  },
);

// ── FEEDBACK ──────────────────────────────────────────────
app.post("/api/feedback", async (req, res) => {
  const { name, content, rating, user_id } = req.body;
  if (!content)
    return res.status(400).json({ error: "Nội dung không được trống" });
  await pool.query(
    `INSERT INTO feedback (user_id,name,content,rating) VALUES ($1,$2,$3,$4)`,
    [user_id || null, name || "Khách hàng", content, rating || 5],
  );
  res.status(201).json({ message: "Cảm ơn bạn đã gửi phản hồi!" });
});

app.get("/api/feedback", auth, can("admin", "manager"), async (req, res) => {
  const r = await pool.query(
    `SELECT f.*, u.username FROM feedback f
     LEFT JOIN users u ON f.user_id=u.id
     ORDER BY f.created_at DESC LIMIT 100`,
  );
  res.json(r.rows);
});

app.put(
  "/api/feedback/:id/status",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    await pool.query(`UPDATE feedback SET status=$1 WHERE id=$2`, [
      req.body.status,
      req.params.id,
    ]);
    res.json({ message: "Cập nhật trạng thái thành công" });
  },
);

// ── MY STATS (cashier) ────────────────────────────────────
app.get("/api/my-stats", auth, async (req, res) => {
  const today = new Date().toISOString().split("T")[0];
  const [orders, att] = await Promise.all([
    pool.query(
      `SELECT COUNT(*) AS cnt, COALESCE(SUM(total),0) AS rev
       FROM orders WHERE cashier_id=$1 AND DATE(created_at)=$2`,
      [req.user.id, today],
    ),
    pool.query(
      `SELECT checkin,checkout FROM attendance WHERE user_id=$1 AND date=$2`,
      [req.user.id, today],
    ),
  ]);
  const att_row = att.rows[0] || {};
  let hours = 0;
  if (att_row.checkin && att_row.checkout) {
    hours = (
      (new Date(att_row.checkout) - new Date(att_row.checkin)) /
      3600000
    ).toFixed(1);
  }
  res.json({
    orders_today: parseInt(orders.rows[0].cnt),
    revenue_today: parseInt(orders.rows[0].rev),
    checkin: att_row.checkin,
    checkout: att_row.checkout,
    hours_worked: hours,
  });
});

// ── KẾT THÚC PHẦN 3/4 ──
// → Tiếp tục paste PHẦN 4/4 ngay bên dưới
// ── PHẦN 4/4: Stats + Reports + Admin Summary + Start ─────

// ── STATS DASHBOARD ───────────────────────────────────────
app.get("/api/stats", auth, can("admin", "manager"), async (req, res) => {
  const today = new Date().toISOString().split("T")[0];
  const [rev, ord, cust, low] = await Promise.all([
    pool.query(
      `SELECT COALESCE(SUM(total),0) AS v FROM orders WHERE DATE(created_at)=$1 AND status='done'`,
      [today],
    ),
    pool.query(`SELECT COUNT(*) AS v FROM orders WHERE DATE(created_at)=$1`, [
      today,
    ]),
    pool.query(`SELECT COUNT(*) AS v FROM users WHERE role_id=7`),
    pool.query(
      `SELECT COUNT(*) AS v FROM products WHERE stock<=min_stock AND is_active=true`,
    ),
  ]);
  const days = ["CN", "T2", "T3", "T4", "T5", "T6", "T7"];
  const weekly = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const ds = d.toISOString().split("T")[0];
    const r = await pool.query(
      `SELECT COALESCE(SUM(total),0) AS v FROM orders WHERE DATE(created_at)=$1 AND status='done'`,
      [ds],
    );
    weekly.push({ label: days[d.getDay()], rev: parseInt(r.rows[0].v) });
  }
  res.json({
    revenue_today: parseInt(rev.rows[0].v),
    orders_today: parseInt(ord.rows[0].v),
    customers: parseInt(cust.rows[0].v),
    low_stock: parseInt(low.rows[0].v),
    today_revenue: parseInt(rev.rows[0].v),
    today_orders: parseInt(ord.rows[0].v),
    total_customers: parseInt(cust.rows[0].v),
    weekly,
  });
});

// ── ADMIN SUMMARY ──────────────────────────────────────────
app.get(
  "/api/admin/summary",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    try {
      const [
        users,
        prods,
        orders,
        lowStock,
        pendingDeliv,
        pendingFeedback,
        vouchers,
      ] = await Promise.all([
        pool.query(`SELECT COUNT(*) AS v FROM users WHERE is_active=true`),
        pool.query(`SELECT COUNT(*) AS v FROM products WHERE is_active=true`),
        pool.query(
          `SELECT COUNT(*) AS v, COALESCE(SUM(total),0) AS rev FROM orders WHERE status='done'`,
        ),
        pool.query(
          `SELECT COUNT(*) AS v FROM products WHERE stock<=min_stock AND is_active=true`,
        ),
        pool.query(
          `SELECT COUNT(*) AS v FROM orders WHERE delivery_status='pending'`,
        ),
        pool.query(`SELECT COUNT(*) AS v FROM feedback WHERE status='pending'`),
        pool.query(`SELECT COUNT(*) AS v FROM vouchers WHERE is_active=true`),
      ]);
      res.json({
        total_users: parseInt(users.rows[0].v),
        total_products: parseInt(prods.rows[0].v),
        total_orders: parseInt(orders.rows[0].v),
        total_revenue: parseInt(orders.rows[0].rev),
        low_stock: parseInt(lowStock.rows[0].v),
        pending_delivery: parseInt(pendingDeliv.rows[0].v),
        pending_feedback: parseInt(pendingFeedback.rows[0].v),
        active_vouchers: parseInt(vouchers.rows[0].v),
      });
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: "Lỗi hệ thống" });
    }
  },
);

// ── REPORTS ───────────────────────────────────────────────

// Báo cáo doanh thu theo ngày
app.get(
  "/api/reports/revenue",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { days = 30 } = req.query;
    const r = await pool.query(
      `SELECT DATE(created_at) AS date,
            COUNT(*) AS order_count,
            COALESCE(SUM(total),0) AS revenue,
            COALESCE(SUM(discount),0) AS total_discount
     FROM orders
     WHERE created_at >= NOW() - INTERVAL '${parseInt(days)} days'
       AND status='done'
     GROUP BY DATE(created_at)
     ORDER BY date DESC`,
    );
    res.json(r.rows);
  },
);

// Báo cáo top sản phẩm bán chạy
app.get(
  "/api/reports/top-products",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { limit = 10 } = req.query;
    const r = await pool.query(
      `SELECT oi.product_name,
            COALESCE(oi.product_id, 0) AS product_id,
            SUM(oi.qty) AS total_qty,
            SUM(oi.qty * oi.price) AS total_revenue
     FROM order_items oi
     JOIN orders o ON oi.order_id=o.id
     WHERE o.status='done'
     GROUP BY oi.product_name, oi.product_id
     ORDER BY total_qty DESC
     LIMIT $1`,
      [parseInt(limit)],
    );
    res.json(r.rows);
  },
);

// Báo cáo theo danh mục
app.get(
  "/api/reports/categories",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const r = await pool.query(
      `SELECT p.category,
            COUNT(DISTINCT p.id) AS product_count,
            COALESCE(SUM(oi.qty),0) AS total_sold,
            COALESCE(SUM(oi.qty * oi.price),0) AS revenue
     FROM products p
     LEFT JOIN order_items oi ON p.id=oi.product_id
     LEFT JOIN orders o ON oi.order_id=o.id AND o.status='done'
     WHERE p.is_active=true
     GROUP BY p.category
     ORDER BY revenue DESC`,
    );
    res.json(r.rows);
  },
);

// Báo cáo nhân viên
app.get(
  "/api/reports/staff-stats",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { month } = req.query;
    const target = month || new Date().toISOString().slice(0, 7);
    const r = await pool.query(
      `SELECT u.id, u.full_name, u.username, ro.display_name AS role_display,
            COUNT(DISTINCT a.date) AS days_worked,
            COUNT(DISTINCT o.id) AS orders_handled,
            COALESCE(SUM(o.total),0) AS revenue_handled,
            COALESCE(s.base_salary,0) AS base_salary,
            COALESCE(s.bonus,0) AS bonus,
            COALESCE(s.deduction,0) AS deduction
     FROM users u
     JOIN roles ro ON u.role_id=ro.id
     LEFT JOIN attendance a ON u.id=a.user_id
       AND TO_CHAR(a.date,'YYYY-MM')=$1
     LEFT JOIN orders o ON u.id=o.cashier_id
       AND TO_CHAR(o.created_at,'YYYY-MM')=$1
       AND o.status='done'
     LEFT JOIN salary s ON u.id=s.user_id
     WHERE u.is_active=true AND ro.name != 'customer'
     GROUP BY u.id, u.full_name, u.username, ro.display_name,
              s.base_salary, s.bonus, s.deduction
     ORDER BY revenue_handled DESC`,
      [target],
    );
    res.json(r.rows);
  },
);

// Xuất báo cáo CSV
app.get(
  "/api/reports/export",
  auth,
  can("admin", "manager"),
  async (req, res) => {
    const { type = "revenue", days = 30 } = req.query;
    let r, headers, rows;

    if (type === "revenue") {
      r = await pool.query(
        `SELECT DATE(created_at) AS date, COUNT(*) AS orders,
              COALESCE(SUM(total),0) AS revenue
       FROM orders WHERE created_at >= NOW() - INTERVAL '${parseInt(days)} days'
         AND status='done'
       GROUP BY DATE(created_at) ORDER BY date DESC`,
      );
      headers = "Ngày,Số đơn,Doanh thu\n";
      rows = r.rows.map((r) => `${r.date},${r.orders},${r.revenue}`).join("\n");
    } else if (type === "products") {
      r = await pool.query(
        `SELECT oi.product_name, SUM(oi.qty) AS qty,
              SUM(oi.qty*oi.price) AS revenue
       FROM order_items oi JOIN orders o ON oi.order_id=o.id
       WHERE o.status='done'
       GROUP BY oi.product_name ORDER BY qty DESC LIMIT 50`,
      );
      headers = "Sản phẩm,Số lượng bán,Doanh thu\n";
      rows = r.rows
        .map((r) => `"${r.product_name}",${r.qty},${r.revenue}`)
        .join("\n");
    } else {
      headers = "Không có dữ liệu\n";
      rows = "";
    }

    const csv = "\uFEFF" + headers + rows; // BOM UTF-8 cho Excel
    res.setHeader("Content-Type", "text/csv; charset=utf-8");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename=baocao_${type}_${Date.now()}.csv`,
    );
    res.send(csv);
  },
);

// ── CUSTOMERS ─────────────────────────────────────────────
app.get(
  "/api/customers",
  auth,
  can("admin", "manager", "cashier"),
  async (req, res) => {
    const r = await pool.query(
      `SELECT u.id, u.username, u.full_name, u.email, u.phone,
            u.created_at, cp.rank, cp.points, cp.total_spent
     FROM users u
     JOIN roles ro ON u.role_id=ro.id
     LEFT JOIN customer_profiles cp ON u.id=cp.user_id
     WHERE ro.name='customer'
     ORDER BY cp.total_spent DESC NULLS LAST`,
    );
    res.json(r.rows);
  },
);

// ── HEALTH CHECK ──────────────────────────────────────────
app.get("/api/health", async (req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ status: "ok", time: new Date().toISOString() });
  } catch (e) {
    res.status(500).json({ status: "error", message: e.message });
  }
});

// ── SERVE FRONTEND ────────────────────────────────────────
app.get("/app.html", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "app.html"));
});
// ✅ Cách 1 - dùng tên tham số rõ ràng (khuyên dùng)
app.get("/{*path}", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ── START SERVER ──────────────────────────────────────────
initDB()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`🚀 MiciMart chạy tại http://localhost:${PORT}`);
      console.log(`📦 Môi trường: ${process.env.NODE_ENV || "development"}`);
    });
  })
  .catch((e) => {
    console.error("❌ Lỗi khởi động:", e.message);
    process.exit(1);
  });
//--```

//---

//## ✅ Hoàn thành! Tóm tắt 4 phần:

//| Phần | Nội dung |
//|------|----------|
//| **1/4** | Setup + DB schema + Seed + Auth (register/login) |
//| **2/4** | Users + Products + Import Excel + Orders + Delivery |
//| **3/4** | Vouchers + Attendance + Salary + Feedback + My Stats |
//| **4/4** | Stats + Admin Summary + Reports 4 loại + Export CSV + Start |

//---

//## Cách ghép file:

//Tạo file `server.js` → paste **4 phần nối tiếp nhau theo thứ tự** → `Ctrl+S` → chạy:
//```
//npm install
//node server.js
//```

//Phải thấy:
//```
//✅ Database sẵn sàng
//🔑 Tài khoản admin: admin12 / admin12
//🚀 MiciMart chạy tại http://localhost:3000 --
