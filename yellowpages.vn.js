const axios = require("axios");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");
require("dotenv").config();

// Hàm gửi yêu cầu HTTP với retry khi gặp lỗi 429
async function fetchWithRetry(url, maxRetries = 5, baseDelay = 1000) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await axios.get(url);
      return response;
    } catch (error) {
      if (error.response && error.response.status === 429) {
        if (attempt === maxRetries) {
          console.error(`Hết số lần thử (${maxRetries}) cho URL: ${url}`);
          throw error;
        }
        const delay = baseDelay * Math.pow(2, attempt - 1);
        console.warn(
          `Lỗi 429 (Too Many Requests) tại URL: ${url}. Thử lại lần ${attempt}/${maxRetries} sau ${delay}ms...`
        );
        await new Promise((resolve) => setTimeout(resolve, delay));
      } else {
        throw error;
      }
    }
  }
}

// Hàm lấy danh sách công ty từ một trang
async function crawlCompanyData(page) {
  try {
    const url = `https://www.yellowpages.vn/cls/224510/van-phong-pham-cong-ty-van-phong-pham.html?page=${page}`;
    const response = await fetchWithRetry(url);
    const $ = cheerio.load(response.data);

    const companyData = [];
    // Tìm tất cả các khối công ty
    $("div.rounded-4.border.bg-white.shadow-sm.mb-3.pb-4").each(
      (index, element) => {
        const name = $(element)
          .find("h2.fs-5.pb-0.text-capitalize a")
          .text()
          .trim();
        const description = $(element)
          .find("div.mt-3.rounded-4.pb-2.h-auto.text_quangcao small")
          .text()
          .trim()
          .replace(/\n/g, " "); // Thay dòng mới bằng dấu cách
        const address = $(element)
          .find(
            "div.float-end.yp_diachi_logo p:has(i.fa.fa-solid.fa-location-arrow) small"
          )
          .text()
          .trim();
        const phone = $(element)
          .find("div.float-end.yp_diachi_logo p:has(i.fa.fa-solid.fa-phone) a")
          .first()
          .text()
          .trim();
        const mobile = $(element)
          .find(
            "div.float-end.yp_diachi_logo p:has(i.fa.fa-solid.fa-mobile-screen-button) a"
          )
          .first()
          .text()
          .trim();
        const email = $(element)
          .find(
            "div.float-end.yp_diachi_logo p:has(i.fa.fa-regular.fa-envelope) a"
          )
          .text()
          .trim();
        const website = $(element)
          .find("div.float-end.yp_diachi_logo p:has(i.fa.fa-solid.fa-globe) a")
          .text()
          .trim();

        // Kiểm tra xem name hoặc description có chứa "sản xuất" không
        const isProductionRelated =
          /sản xuất/i.test(name) || /sản xuất/i.test(description);

        if (name && isProductionRelated) {
          companyData.push({
            name,
            description,
            address,
            phone,
            mobile,
            email,
            website,
          });
        }
      }
    );

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
    { header: "Mô tả", key: "description", width: 100 },
    { header: "Địa chỉ", key: "address", width: 50 },
    { header: "Điện thoại", key: "phone", width: 20 },
    { header: "Di động/Hotline", key: "mobile", width: 20 },
    { header: "Email", key: "email", width: 30 },
    { header: "Website", key: "website", width: 30 },
  ];

  // Thêm dữ liệu
  companyData.forEach((company, index) => {
    worksheet.addRow({
      index: index + 1,
      name: company.name,
      description: company.description,
      address: company.address,
      phone: company.phone,
      mobile: company.mobile,
      email: company.email,
      website: company.website,
    });
  });

  // Lưu file
  const fileName = `yellowpages.xlsx`;
  await workbook.xlsx.writeFile(fileName);
  console.log(`Đã xuất danh sách công ty vào file: ${fileName}`);
}

// Hàm chính để crawl và xử lý
async function main() {
  let allCompanyData = [];

  // Số trang tối đa là 230
  const maxPages = 230;

  console.log(`Ngành nghề: Văn phòng phẩm - Công ty sản xuất`);

  // Lặp qua các trang từ 1 đến maxPages
  for (let page = 1; page <= maxPages; page++) {
    console.log(`Đang crawl trang ${page}...`);
    const companyData = await crawlCompanyData(page);
    allCompanyData.push(...companyData);
    console.log(`Hoàn thành trang ${page}. Số công ty: ${companyData.length}`);
    if (companyData.length === 0) {
      console.log(`Không tìm thấy dữ liệu ở trang ${page}. Dừng crawl.`);
      // break;
    }

    // Delay 100ms giữa các request để tránh bị chặn
    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  // Loại bỏ trùng lặp dựa trên tên công ty
  const uniqueCompanyData = [];
  const seenNames = new Set();

  for (const company of allCompanyData) {
    if (!seenNames.has(company.name)) {
      seenNames.add(company.name);
      uniqueCompanyData.push(company);
    }
  }

  // In ra danh sách công ty
  console.log("\nDanh sách công ty không trùng lặp:");
  uniqueCompanyData.forEach((company, index) => {
    console.log(`${index + 1}. ${company.name}`);
    console.log(`   Mô tả: ${company.description}`);
    console.log(`   Địa chỉ: ${company.address}`);
    console.log(`   Điện thoại: ${company.phone}`);
    console.log(`   Di động/Hotline: ${company.mobile}`);
    console.log(`   Email: ${company.email}`);
    console.log(`   Website: ${company.website}`);
    console.log("");
  });
  console.log(`\nTổng số công ty không trùng lặp: ${uniqueCompanyData.length}`);

  // Xuất vào file Excel
  await exportToExcel(uniqueCompanyData);
}

// Chạy chương trình
main().catch((error) => {
  console.error("Lỗi trong quá trình chạy chương trình:", error.message);
});
