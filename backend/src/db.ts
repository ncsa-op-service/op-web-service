import "dotenv/config";
import { Pool } from "pg";

if (!process.env.DATABASE_URL) {
  throw new Error("ไม่พบ DATABASE_URL");
}

export const pool = new Pool({
  connectionString:
    process.env.DATABASE_URL,
});

export async function initializeDatabase():
  Promise<void> {
  // =========================================================
  // ประวัติการอัปโหลดไฟล์
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS uploads (
      id SERIAL PRIMARY KEY,
      original_file_name VARCHAR(255) NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `);

  // =========================================================
  // LINE URL
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS line_links (
      id SERIAL PRIMARY KEY,

      upload_id INTEGER
        REFERENCES uploads(id)
        ON DELETE CASCADE,

      row_number INTEGER NOT NULL,
      name VARCHAR(255),
      original_url TEXT NOT NULL,
      final_url TEXT,

      status VARCHAR(30)
        NOT NULL
        DEFAULT 'WAITING',

      message TEXT,

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      checked_at TIMESTAMPTZ
    );
  `);

  // =========================================================
  // ผู้ใช้งานระบบ
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      name VARCHAR(100) NOT NULL,
      email VARCHAR(150) UNIQUE NOT NULL,
      password_hash VARCHAR(255) NOT NULL,

      role VARCHAR(20)
        NOT NULL
        DEFAULT 'viewer'
        CHECK (
          role IN (
            'super_admin',
            'editor',
            'viewer'
          )
        ),

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  // =========================================================
  // System Log
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS admin_logs (
      id SERIAL PRIMARY KEY,

      admin_id INTEGER
        REFERENCES users(id)
        ON DELETE SET NULL,

      action_type VARCHAR(50) NOT NULL,
      message TEXT NOT NULL,
      ip_address VARCHAR(100),

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  // =========================================================
  // URL Fake Web - Batch
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS url_fake_batches (
      id SERIAL PRIMARY KEY,

      original_file_name VARCHAR(255),
      selected_provider VARCHAR(50),
      public_ip VARCHAR(100),
      isp TEXT,
      detected_provider VARCHAR(50),

      case_type VARCHAR(30)
        NOT NULL
        DEFAULT 'fake_domain',

      detected_date DATE
        NOT NULL
        DEFAULT CURRENT_DATE,

      case_status VARCHAR(20)
        NOT NULL
        DEFAULT 'active',

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  /*
    เพิ่มคอลัมน์ให้ฐานข้อมูลเดิม

    CREATE TABLE IF NOT EXISTS
    จะไม่เพิ่มคอลัมน์ให้ตารางที่มีอยู่แล้ว
  */
  await pool.query(`
    ALTER TABLE url_fake_batches
      ADD COLUMN IF NOT EXISTS
        case_type VARCHAR(30),

      ADD COLUMN IF NOT EXISTS
        detected_date DATE,

      ADD COLUMN IF NOT EXISTS
        case_status VARCHAR(20)
        DEFAULT 'active';
  `);

  // เติมข้อมูลให้รายการเก่า
  await pool.query(`
    UPDATE url_fake_batches
    SET
      case_type = COALESCE(
        case_type,
        'fake_domain'
      ),

      detected_date = COALESCE(
        detected_date,
        created_at::date
      ),

      case_status = COALESCE(
        case_status,
        'active'
      )
    WHERE
      case_type IS NULL
      OR detected_date IS NULL
      OR case_status IS NULL;
  `);

  // บังคับให้ข้อมูลใหม่ต้องมีค่า
  await pool.query(`
    ALTER TABLE url_fake_batches
      ALTER COLUMN case_type
        SET DEFAULT 'fake_domain',

      ALTER COLUMN case_type
        SET NOT NULL,

      ALTER COLUMN detected_date
        SET DEFAULT CURRENT_DATE,

      ALTER COLUMN detected_date
        SET NOT NULL,

      ALTER COLUMN case_status
        SET DEFAULT 'active',

      ALTER COLUMN case_status
        SET NOT NULL;
  `);

  // ตรวจสอบค่าประเภทและสถานะ
  await pool.query(`
    DO $$
    BEGIN
      IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conname =
          'url_fake_batches_case_type_check'
      ) THEN
        ALTER TABLE url_fake_batches
          ADD CONSTRAINT
            url_fake_batches_case_type_check
          CHECK (
            case_type IN (
              'fake_domain',
              'fake_web',
              'c2_oss',
              'fake_line'
            )
          );
      END IF;

      IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conname =
          'url_fake_batches_case_status_check'
      ) THEN
        ALTER TABLE url_fake_batches
          ADD CONSTRAINT
            url_fake_batches_case_status_check
          CHECK (
            case_status IN (
              'active',
              'inactive'
            )
          );
      END IF;
    END
    $$;
  `);

  // =========================================================
  // URL Fake Web - Results
  // =========================================================
  await pool.query(`
    CREATE TABLE IF NOT EXISTS url_fake_results (
      id SERIAL PRIMARY KEY,

      batch_id INTEGER NOT NULL
        REFERENCES url_fake_batches(id)
        ON DELETE CASCADE,

      case_id VARCHAR(100) NOT NULL,
      url_sms TEXT NOT NULL,

      ais_result TEXT,
      true_dtac_result TEXT,
      nt_result TEXT,
      cloudflare_result TEXT,

      ais_status VARCHAR(20),
      true_dtac_status VARCHAR(20),
      nt_status VARCHAR(20),
      cloudflare_status VARCHAR(20),

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  /*
    สำหรับฐานข้อมูลเดิมที่มีตาราง url_fake_results อยู่แล้ว
    CREATE TABLE IF NOT EXISTS จะไม่เพิ่มคอลัมน์ใหม่ให้
    จึงต้อง ALTER TABLE เพิ่ม status แยกอีกครั้ง
  */
  await pool.query(`
    ALTER TABLE url_fake_results
      ADD COLUMN IF NOT EXISTS
        ais_status VARCHAR(20),

      ADD COLUMN IF NOT EXISTS
        true_dtac_status VARCHAR(20),

      ADD COLUMN IF NOT EXISTS
        nt_status VARCHAR(20),

      ADD COLUMN IF NOT EXISTS
        cloudflare_status VARCHAR(20);
  `);

  // =========================================================
  // SSL / TLS CHECKER
  // =========================================================

  /*
    1 งานตรวจ = 1 Job
    เช่น monitored-domains-unique.txt จำนวน 308,620 URL

    status:
      READY      = พร้อมตรวจ
      RUNNING    = กำลังตรวจ
      PAUSED     = หยุดพัก
      COMPLETED  = ตรวจครบทั้งหมด
      ERROR      = งานมีข้อผิดพลาด
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS ssl_check_jobs (
      id SERIAL PRIMARY KEY,

      original_file_name VARCHAR(255)
        NOT NULL,

      total_urls INTEGER
        NOT NULL
        CHECK (total_urls >= 0),

      round_size INTEGER
        NOT NULL
        DEFAULT 500
        CHECK (round_size > 0),

      request_size INTEGER
        NOT NULL
        DEFAULT 100
        CHECK (request_size > 0),

      total_rounds INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (total_rounds >= 0),

      checked_count INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (checked_count >= 0),

      completed_rounds INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (completed_rounds >= 0),

      current_round INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (current_round >= 0),

      status VARCHAR(20)
        NOT NULL
        DEFAULT 'READY'
        CHECK (
          status IN (
            'READY',
            'RUNNING',
            'PAUSED',
            'COMPLETED',
            'ERROR'
          )
        ),

      created_by INTEGER
        REFERENCES users(id)
        ON DELETE SET NULL,

      last_error TEXT,

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      started_at TIMESTAMPTZ,

      paused_at TIMESTAMPTZ,

      completed_at TIMESTAMPTZ
    );
  `);

  /*
    เก็บ URL ต้นฉบับทุกตัวตามลำดับในไฟล์

    สำคัญมาก:
    ถ้าผู้ใช้ปิด Browser แล้วกลับมาใหม่
    Backend ยังรู้ว่า URL ลำดับไหนต้องตรวจต่อ
    โดยไม่ต้องอัปโหลดไฟล์เดิมใหม่
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS ssl_check_items (
      id BIGSERIAL PRIMARY KEY,

      job_id INTEGER NOT NULL
        REFERENCES ssl_check_jobs(id)
        ON DELETE CASCADE,

      file_order INTEGER NOT NULL
        CHECK (file_order > 0),

      round_number INTEGER NOT NULL
        CHECK (round_number > 0),

      url TEXT NOT NULL,

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      UNIQUE (
        job_id,
        file_order
      )
    );
  `);

  /*
    สถานะของแต่ละรอบ

    1 รอบ = 500 URL
    เช่น:
      รอบ 1 = 1 - 500
      รอบ 2 = 501 - 1000
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS ssl_check_rounds (
      id SERIAL PRIMARY KEY,

      job_id INTEGER NOT NULL
        REFERENCES ssl_check_jobs(id)
        ON DELETE CASCADE,

      round_number INTEGER NOT NULL
        CHECK (round_number > 0),

      start_order INTEGER NOT NULL
        CHECK (start_order > 0),

      end_order INTEGER NOT NULL
        CHECK (end_order >= start_order),

      total_urls INTEGER NOT NULL
        CHECK (total_urls >= 0),

      checked_count INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (checked_count >= 0),

      status VARCHAR(20)
        NOT NULL
        DEFAULT 'WAITING'
        CHECK (
          status IN (
            'WAITING',
            'RUNNING',
            'PAUSED',
            'COMPLETED',
            'ERROR'
          )
        ),

      started_at TIMESTAMPTZ,

      paused_at TIMESTAMPTZ,

      completed_at TIMESTAMPTZ,

      updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      UNIQUE (
        job_id,
        round_number
      )
    );
  `);

  /*
    ผลตรวจ SSL/TLS จริง

    UNIQUE(job_id, file_order)
    ทำให้ URL ลำดับเดิมมีผลได้เพียง 1 แถว
    และสามารถใช้ UPSERT ตอนตรวจซ้ำได้
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS ssl_check_results (
      id BIGSERIAL PRIMARY KEY,

      job_id INTEGER NOT NULL
        REFERENCES ssl_check_jobs(id)
        ON DELETE CASCADE,

      round_id INTEGER NOT NULL
        REFERENCES ssl_check_rounds(id)
        ON DELETE CASCADE,

      file_order INTEGER NOT NULL
        CHECK (file_order > 0),

      url TEXT NOT NULL,

      hostname TEXT,

      port INTEGER,

      valid_from TIMESTAMPTZ,

      expiration_date TIMESTAMPTZ,

      days_left INTEGER,

      status VARCHAR(30)
        NOT NULL
        CHECK (
          status IN (
            'VALID',
            'EXPIRING_SOON',
            'EXPIRED',
            'ERROR'
          )
        ),

      issuer TEXT,

      subject TEXT,

      checked_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      error TEXT,

      updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      UNIQUE (
        job_id,
        file_order
      )
    );
  `);


  // =========================================================
  // CLAB - Snapshot
  // =========================================================

  /*
    1 ครั้งที่กด "บันทึก CLAB" = 1 Snapshot

    เก็บชื่อไฟล์ต้นฉบับและผู้ที่กดบันทึก
    Super Admin / Editor เป็นผู้บันทึก
    Viewer ใช้สำหรับอ่าน Snapshot ล่าสุด
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS clab_snapshots (
      id SERIAL PRIMARY KEY,

      original_file_name VARCHAR(255),

      saved_by INTEGER
        REFERENCES users(id)
        ON DELETE SET NULL,

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  /*
    แต่ละ Snapshot มีได้หลาย Sheet

    เก็บข้อมูลแต่ละ Sheet เป็น JSONB เพราะโครงสร้างของ
    Overview VA / Report / OP02 / OP03 /
    Waiting Organizations / Email Service ไม่เหมือนกัน

    rows จะเก็บข้อมูลแบบ array-of-arrays จาก Excel
    เพื่อสามารถสร้าง Worksheet กลับมาใช้กับ Component เดิมได้
  */
  await pool.query(`
    CREATE TABLE IF NOT EXISTS clab_snapshot_sheets (
      id SERIAL PRIMARY KEY,

      snapshot_id INTEGER NOT NULL
        REFERENCES clab_snapshots(id)
        ON DELETE CASCADE,

      original_sheet_name VARCHAR(255)
        NOT NULL,

      detected_type VARCHAR(50)
        NOT NULL,

      detected_label VARCHAR(120)
        NOT NULL,

      confidence INTEGER
        NOT NULL
        DEFAULT 0
        CHECK (
          confidence >= 0
          AND confidence <= 100
        ),

      headers JSONB
        NOT NULL
        DEFAULT '[]'::jsonb,

      rows JSONB
        NOT NULL
        DEFAULT '[]'::jsonb,

      /*
        เก็บขอบเขต Sheet และ Merge Cell
        เพื่อสร้าง Worksheet กลับมาได้ตรงโครงเดิม
      */
      sheet_ref TEXT,

      merges JSONB
        NOT NULL
        DEFAULT '[]'::jsonb,

      created_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW()
    );
  `);

  /*
    สำหรับฐานข้อมูลเดิมที่มีตาราง clab_snapshot_sheets อยู่แล้ว
    CREATE TABLE IF NOT EXISTS จะไม่เพิ่มคอลัมน์ใหม่ให้
    จึงต้อง ALTER TABLE เพิ่ม updated_at แยกอีกครั้ง
  */
  await pool.query(`
    ALTER TABLE clab_snapshot_sheets
      ADD COLUMN IF NOT EXISTS
        updated_at TIMESTAMPTZ
        NOT NULL
        DEFAULT NOW(),

      ADD COLUMN IF NOT EXISTS
        sheet_ref TEXT,

      ADD COLUMN IF NOT EXISTS
        merges JSONB
        NOT NULL
        DEFAULT '[]'::jsonb;
  `);

  // =========================================================
  // CLAB - Indexes
  // =========================================================

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_clab_snapshots_created_at
    ON clab_snapshots (
      created_at DESC,
      id DESC
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_clab_snapshot_sheets_snapshot
    ON clab_snapshot_sheets (
      snapshot_id,
      id
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_clab_snapshot_sheets_type_updated
    ON clab_snapshot_sheets (
      detected_type,
      updated_at DESC,
      id DESC
    );
  `);

  // =========================================================
  // INDEXES
  // =========================================================

  // Main Summary - URL Fake Web
  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_url_fake_batches_summary
    ON url_fake_batches (
      case_type,
      detected_date DESC,
      id DESC
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_url_fake_results_batch_url
    ON url_fake_results (
      batch_id,
      url_sms
    );
  `);

  // SSL Job
  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_jobs_created_at
    ON ssl_check_jobs (
      created_at DESC,
      id DESC
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_jobs_status
    ON ssl_check_jobs (
      status,
      updated_at DESC
    );
  `);

  // SSL Items
  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_items_job_round_order
    ON ssl_check_items (
      job_id,
      round_number,
      file_order
    );
  `);

  // SSL Rounds
  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_rounds_job_round
    ON ssl_check_rounds (
      job_id,
      round_number
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_rounds_job_status
    ON ssl_check_rounds (
      job_id,
      status
    );
  `);

  // SSL Results
  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_results_job_round_order
    ON ssl_check_results (
      job_id,
      round_id,
      file_order
    );
  `);

  await pool.query(`
    CREATE INDEX IF NOT EXISTS
      idx_ssl_check_results_job_status
    ON ssl_check_results (
      job_id,
      status
    );
  `);

  console.log(
    "Database initialized successfully"
  );
}
