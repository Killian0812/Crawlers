const axios = require("axios");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");
require("dotenv").config();

// Hàm lấy danh sách công ty từ một trang
async function crawlCompanyData(page) {
  try {
    const url = `https://thuvienphapluat.vn/ma-so-thue/tra-cuu-ma-so-thue-doanh-nghiep?timtheo=ma-so-thue&tukhoa=&ngaycaptu=&ngaycapden=&ngaydongmsttu=&ngaydongmstden=&vondieuletu=&vondieuleden=&loaihinh=0&nganhnghe=${process.env.CODE}&tinhthanhpho=0&quanhuyen=0&phuongxa=0&page=${page}`;
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);

    const companyData = [];
    // Tìm tất cả các hàng trong bảng
    $("table.table-bordered tbody tr").each((index, element) => {
      const taxCode = $(element)
        .find('td:nth-child(2) a[title*="Mã số thuế"]')
        .text()
        .trim();
      const companyName = $(element)
        .find('td:nth-child(3) a[style*="font-weight:bold"]')
        .text()
        .trim();

      if (companyName && taxCode) {
        companyData.push({ name: companyName, taxCode });
      }
    });

    return companyData;
  } catch (error) {
    console.error(`Lỗi khi crawl trang ${page}:`, error.message);
    return [];
  }
}

// Hàm xuất danh sách công ty vào file Excel
async function exportToExcel(companyData) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Danh sách công ty");

  // Thiết lập tiêu đề cột
  worksheet.columns = [
    { header: "STT", key: "index", width: 10 },
    { header: "Tên công ty", key: "name", width: 50 },
    { header: "Mã số thuế", key: "taxCode", width: 20 },
  ];

  // Thêm dữ liệu
  companyData.forEach((company, index) => {
    worksheet.addRow({
      index: index + 1,
      name: company.name,
      taxCode: company.taxCode,
    });
  });

  // Lưu file
  const code = process.env.CODE;
  const fileName = `${code}-${process.env.START}.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log(`Đã xuất danh sách công ty vào file: ${fileName}`);
}

// Hàm chính để crawl và xử lý
async function main() {
  let allCompanyData = [];

  // Lấy số trang từ env, tối đa 50
  const maxPages = 500;
  if (isNaN(maxPages) || maxPages <= 0) {
    throw new Error("PAGE phải là số hợp lệ từ 1 đến 50");
  }

  console.log(`${process.env.CODE}`);

  // Lặp qua các trang từ 1 đến maxPages
  for (let page = process.env.START; page <= maxPages; page++) {
    console.log(`Đang crawl trang ${page}...`);
    const companyData = await crawlCompanyData(page);
    allCompanyData.push(...companyData);
    console.log(`Hoàn thành trang ${page}. Số công ty: ${companyData.length}`);
    if (companyData.length === 0) break;

    // Delay 1 giây giữa các request để tránh bị chặn
    await new Promise((resolve) => setTimeout(resolve, 10));
  }

  // Loại bỏ trùng lặp dựa trên mã số thuế
  const uniqueCompanyData = [];
  const seenTaxCodes = new Set();

  for (const company of allCompanyData) {
    if (!seenTaxCodes.has(company.taxCode)) {
      seenTaxCodes.add(company.taxCode);
      uniqueCompanyData.push(company);
    }
  }

  // In ra danh sách công ty
  console.log("\nDanh sách công ty không trùng lặp:");
  uniqueCompanyData.forEach((company, index) => {
    console.log(`${index + 1}. ${company.name} - ${company.taxCode}`);
  });
  console.log(`\nTổng số công ty không trùng lặp: ${uniqueCompanyData.length}`);

  // Xuất vào file Excel
  await exportToExcel(uniqueCompanyData);
}

// Chạy chương trình
main().catch((error) => {
  console.error("Lỗi trong quá trình chạy chương trình:", error.message);
});
