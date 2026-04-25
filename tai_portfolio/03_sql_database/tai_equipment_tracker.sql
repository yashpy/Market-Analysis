-- ============================================================
-- Temple Allen Industries – Equipment Field Tracker
-- TAI Business Development & Data Analyst Portfolio
-- Author: Yadnesh Deshpande
-- ============================================================

-- Customers
CREATE TABLE customers (
    customer_id     SERIAL PRIMARY KEY,
    company_name    VARCHAR(120) NOT NULL,
    industry        VARCHAR(60),   -- Aerospace, Marine, Windpower, Transportation
    country         VARCHAR(60),
    state           VARCHAR(60),
    contact_name    VARCHAR(100),
    contact_email   VARCHAR(120),
    created_at      TIMESTAMP DEFAULT NOW()
);

-- EMMA Units
CREATE TABLE emma_units (
    unit_id         SERIAL PRIMARY KEY,
    serial_number   VARCHAR(50) UNIQUE NOT NULL,
    model_version   VARCHAR(30),
    manufacture_date DATE,
    status          VARCHAR(30) DEFAULT 'In Stock',
    -- Status: In Stock | Deployed | Demo | Under Repair | Returned
    customer_id     INT REFERENCES customers(customer_id),
    deployed_date   DATE,
    notes           TEXT
);

-- Warranties
CREATE TABLE warranties (
    warranty_id     SERIAL PRIMARY KEY,
    unit_id         INT REFERENCES emma_units(unit_id),
    start_date      DATE NOT NULL,
    end_date        DATE NOT NULL,
    warranty_type   VARCHAR(40),  -- Standard | Extended | On-Site
    is_active       BOOLEAN GENERATED ALWAYS AS (CURRENT_DATE <= end_date) STORED
);

-- Consumables Inventory
CREATE TABLE consumables (
    consumable_id   SERIAL PRIMARY KEY,
    item_name       VARCHAR(100) NOT NULL,
    sku             VARCHAR(50),
    unit_price      NUMERIC(10,2),
    stock_qty       INT DEFAULT 0,
    reorder_level   INT DEFAULT 10
);

-- Consumable Orders (per unit/customer)
CREATE TABLE consumable_orders (
    order_id        SERIAL PRIMARY KEY,
    unit_id         INT REFERENCES emma_units(unit_id),
    consumable_id   INT REFERENCES consumables(consumable_id),
    quantity        INT,
    order_date      DATE DEFAULT CURRENT_DATE,
    shipped_date    DATE,
    total_cost      NUMERIC(10,2)
);

-- Proposals & Quotes
CREATE TABLE proposals (
    proposal_id     SERIAL PRIMARY KEY,
    customer_id     INT REFERENCES customers(customer_id),
    created_date    DATE DEFAULT CURRENT_DATE,
    status          VARCHAR(30) DEFAULT 'Draft',
    -- Draft | Sent | Negotiating | Won | Lost
    units_quoted    INT,
    quoted_price    NUMERIC(12,2),
    notes           TEXT
);

-- ============================================================
-- SAMPLE DATA
-- ============================================================

INSERT INTO customers (company_name, industry, country, state, contact_name, contact_email) VALUES
('Boeing MRO Services',     'Aerospace',    'USA', 'WA', 'James Carter',  'jcarter@boeing.com'),
('Lockheed Martin',         'Defense',      'USA', 'MD', 'Sara Mitchell', 'smitchell@lm.com'),
('Vestas Wind Systems',     'Windpower',    'DEN', NULL,  'Lars Nielsen',  'lnielsen@vestas.com'),
('Huntington Ingalls',      'Marine',       'USA', 'VA', 'Tom Reed',      'treed@hii.com'),
('Delta TechOps',           'Aerospace',    'USA', 'GA', 'Amy Chen',      'achen@delta.com');

INSERT INTO emma_units (serial_number, model_version, manufacture_date, status, customer_id, deployed_date) VALUES
('EMMA-2024-001', 'v2.1', '2024-01-15', 'Deployed', 1, '2024-03-01'),
('EMMA-2024-002', 'v2.1', '2024-02-20', 'Deployed', 2, '2024-04-15'),
('EMMA-2024-003', 'v2.2', '2024-05-10', 'Demo',     3, '2024-06-01'),
('EMMA-2024-004', 'v2.2', '2024-07-01', 'Deployed', 4, '2024-08-20'),
('EMMA-2025-001', 'v2.3', '2025-01-10', 'In Stock', NULL, NULL);

INSERT INTO consumables (item_name, sku, unit_price, stock_qty, reorder_level) VALUES
('Sanding Disc 80-grit',    'SD-080', 2.50,  500, 100),
('Sanding Disc 120-grit',   'SD-120', 2.75,  400, 100),
('Sanding Disc 220-grit',   'SD-220', 3.00,  300,  80),
('Backing Pad',             'BP-001', 18.00,  50,  10),
('Dust Collection Filter',  'DCF-02', 45.00,  30,   5);

-- ============================================================
-- USEFUL QUERIES
-- ============================================================

-- 1. All deployed units with customer info
SELECT
    e.serial_number,
    e.model_version,
    c.company_name,
    c.industry,
    e.deployed_date,
    AGE(CURRENT_DATE, e.deployed_date) AS time_deployed
FROM emma_units e
JOIN customers c ON e.customer_id = c.customer_id
WHERE e.status = 'Deployed'
ORDER BY e.deployed_date;

-- 2. Consumables below reorder level
SELECT
    item_name, sku, stock_qty, reorder_level,
    (reorder_level - stock_qty) AS units_to_order
FROM consumables
WHERE stock_qty < reorder_level
ORDER BY (reorder_level - stock_qty) DESC;

-- 3. Revenue by industry segment
SELECT
    c.industry,
    COUNT(DISTINCT e.unit_id)       AS units_deployed,
    SUM(p.quoted_price)             AS total_quoted,
    AVG(p.quoted_price)             AS avg_deal_size
FROM customers c
LEFT JOIN emma_units e  ON e.customer_id = c.customer_id
LEFT JOIN proposals p   ON p.customer_id = c.customer_id AND p.status = 'Won'
GROUP BY c.industry
ORDER BY total_quoted DESC NULLS LAST;

-- 4. Active warranties expiring within 90 days
SELECT
    e.serial_number,
    c.company_name,
    w.end_date,
    w.warranty_type,
    (w.end_date - CURRENT_DATE) AS days_remaining
FROM warranties w
JOIN emma_units e  ON w.unit_id  = e.unit_id
JOIN customers  c  ON e.customer_id = c.customer_id
WHERE w.is_active = TRUE
  AND w.end_date <= CURRENT_DATE + INTERVAL '90 days'
ORDER BY w.end_date;

-- 5. Win rate by proposal status
SELECT
    status,
    COUNT(*)                                    AS count,
    ROUND(COUNT(*) * 100.0 / SUM(COUNT(*)) OVER(), 1) AS pct
FROM proposals
GROUP BY status
ORDER BY count DESC;
