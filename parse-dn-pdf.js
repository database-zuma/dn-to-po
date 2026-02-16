const fs = require('fs');
const { PDFParse } = require('pdf-parse');

/**
 * Parse DN PDF file and extract data in same format as Excel
 * @param {string} pdfPath - Path to PDF file
 * @returns {Object} DN data (dnNumber, date, warehouse, customer, items)
 */
async function parseDNPdf(pdfPath) {
  const parser = new PDFParse({ url: pdfPath });
  const data = await parser.getText();
  const text = data.text;
  
  // Extract DN Number
  const dnMatch = text.match(/Number\s+(DN\/DDD\/[^\s]+)/);
  const dnNumber = dnMatch ? dnMatch[1] : null;
  
  // Extract Date
  const dateMatch = text.match(/Date\s+(\d+\s+\w+\s+\d+)/);
  const dateStr = dateMatch ? dateMatch[1] : null;
  
  // Extract Warehouse
  const warehouseMatch = text.match(/Warehouse\s+([^\n]+)/);
  const warehouse = warehouseMatch ? warehouseMatch[1].trim() : null;
  
  // Extract Customer (Delivery To)
  const customerMatch = text.match(/Delivery To\s+([^\n]+)/);
  const customer = customerMatch ? customerMatch[1].trim() : null;
  
  // Extract items
  const items = [];
  const lines = text.split('\n');
  let inItemSection = false;
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Detect start of item section
    if (line.includes('Item Kode') && line.includes('Name Article')) {
      inItemSection = true;
      continue;
    }
    
    // Detect end of item section
    if (line.includes('Quantity total') || line.includes('Prepared by')) {
      inItemSection = false;
      break;
    }
    
    // Parse item lines
    if (inItemSection && line.length > 0) {
      // Pattern: ITEM_CODE NAME ARTICLE QTY UNIT
      // Example: M1CAV223Z39 MEN CLASSIC 23, 39, NAVY BLACK 1 PAIR
      // Parse from right to left: last is UNIT, second-last is QTY
      const parts = line.split(/\s+/);
      
      if (parts.length >= 4) {
        const unit = parts[parts.length - 1]; // Last: UNIT
        const qtyStr = parts[parts.length - 2]; // Second last: QTY
        const qty = parseInt(qtyStr);
        
        // First part starting with letter+digit is Item Kode
        let itemCodeIdx = -1;
        for (let j = 0; j < parts.length; j++) {
          if (/^[A-Z]\d/.test(parts[j])) {
            itemCodeIdx = j;
            break;
          }
        }
        
        if (itemCodeIdx >= 0 && !isNaN(qty)) {
          const itemCode = parts[itemCodeIdx];
          // Everything between itemCode and QTY is Name Article
          const nameArticle = parts.slice(itemCodeIdx + 1, parts.length - 2).join(' ');
          
          items.push({
            'Item Kode': itemCode,
            'Name Article': nameArticle,
            'Qty': qty,
            'Unit': unit
          });
        }
      }
    }
  }
  
  // Normalize field names to match Excel parser output
  return {
    dnNumber,
    dnDate: dateStr,
    warehouse,
    customerName: customer,
    items: items.map(item => ({
      kode: item['Item Kode'],         // Field name MUST match: kode (not itemKode)
      nameArticle: item['Name Article'],
      qty: item.Qty,
      unit: item.Unit
    })),
    totalQty: items.reduce((sum, item) => sum + item.Qty, 0),
    totalSku: items.length
  };
}

// CLI usage
if (require.main === module) {
  const pdfPath = process.argv[2];
  
  if (!pdfPath) {
    console.error('Usage: node parse-dn-pdf.js <pdf_file>');
    process.exit(1);
  }
  
  parseDNPdf(pdfPath)
    .then(data => {
      console.log(JSON.stringify(data, null, 2));
    })
    .catch(err => {
      console.error('Error:', err.message);
      process.exit(1);
    });
}

module.exports = { parseDNPdf };
