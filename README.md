# Excel Data Exploration – Product Sales Dataset

## 📂 Dataset Overview
This project explores a product sales dataset with the following columns:
- Product ID (includes date + country code)
- Product Name
- Brand Name
- Price ($)
- Quantity
- Category

Countries are embedded in Product ID (US, UK, IN, CA, etc.), and categories include Electronics, Fashion, Kitchen, Outdoor, and Accessories.

---

## 🔧 Excel Techniques Used

### Named Table
- Converted dataset into a structured table for easy reference.
- Benefits: Auto‑filter, dynamic ranges, easier formula writing.

### SUM
- `=SUM(QuantityRange)` → Total quantity of all products.

### SUMIF
- `=SUMIF(CategoryRange,"Electronics",RevenueRange)`  
- Purpose: Category‑wise revenue calculation.

### COUNT & COUNTIF
- `=COUNT(ProductNameRange)` → Total products.  
- `=COUNTIF(CategoryRange,"Fashion")` → Number of fashion items.

### AVERAGE & AVERAGEIF
- `=AVERAGE(PriceRange)` → Average product price.  
- `=AVERAGEIF(CategoryRange,"Kitchen",PriceRange)` → Average price of kitchen items.

### LEFT, MID, RIGHT
- Extracted country code from Product ID.  
- Example: `=RIGHT(A2,2)` → "US", "UK", etc.

### MIN & MAX
- `=MIN(PriceRange)` → Lowest priced product.  
- `=MAX(PriceRange)` → Highest priced product.

---

## 📊 Insights Generated
- **Revenue per Product** = Price × Quantity  
- **Category‑wise comparison** using Pivot Tables  
- **Brand performance** (Nike vs Adidas vs Dell, etc.)  
- **Country‑wise sales** using extracted codes  
- **Top 5 products by revenue** using SORT & FILTER  
