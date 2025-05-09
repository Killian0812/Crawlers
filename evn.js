const axios = require("axios");
const cheerio = require("cheerio");
const XLSX = require("xlsx");

async function fetchPageLinks(page) {
  const url = `https://www.evn.com.vn/vi-VN/news-l/Thong-tin-tom-tat-van-hanh-HTD-Quoc-gia-60-2015?page=${page}`;
  try {
    const response = await axios.get(url);
    const $ = cheerio.load(response.data);
    const links = [];
    $("a")
      .filter((_, element) => $(element).attr("class") === "xanhEVN")
      .each((_, element) => {
        const href = $(element).attr("href");
        const title = $(element).text().trim();
        const date = title.split(" ").pop();
        links.push({ href: `https://www.evn.com.vn${href}`, title: date });
      });
    return links;
  } catch (error) {
    console.error(`Error fetching page ${page}:`, error.message);
    return [];
  }
}

async function fetchTableData(link) {
  try {
    const response = await axios.get(link.href);
    const $ = cheerio.load(response.data);
    const tables = $("table");
    const targetTable = tables.length > 1 ? tables.eq(1) : tables.eq(0);
    const rows = targetTable.find("tbody tr");
    let cells;

    // Check if the table has the new format (with header row containing "CÔNG SUẤT HUY ĐỘNG")
    const firstRowText = $(rows[0]).text();
    if (firstRowText.includes("CÔNG SUẤT HUY ĐỘNG")) {
      // New format: extract the third row ("Toàn Quốc")
      const targetRow = rows.eq(2);
      cells = targetRow
        .find("td")
        .map((_, td) => $(td).text().trim())
        .get();
    } else {
      // Original format: extract the first row
      const targetRow = rows.first();
      cells = targetRow
        .find("td")
        .map((_, td) => $(td).text().trim())
        .get();
    }

    return {
      Ngày: link.title,
      Mục: cells[0] || "",
      "Khi phụ tải vào thấp điểm trưa": cells[1] || "",
      "Khi phụ tải vào cao điểm tối": cells[2] || "",
    };
  } catch (error) {
    console.error(`Error fetching data from ${link.href}:`, error.message);
    return null;
  }
}

async function crawlPages() {
  const allData = [];
  for (let page = 1; page <= 59; page++) {
    console.log(`Crawling page ${page}...`);
    const links = await fetchPageLinks(page);
    for (const link of links) {
      const data = await fetchTableData(link);
      if (data) {
        allData.push(data);
      }
    }
  }
  return allData;
}

function saveToExcel(data) {
  const worksheet = XLSX.utils.json_to_sheet(data, {
    header: [
      "Ngày",
      "Mục",
      "Khi phụ tải vào thấp điểm trưa",
      "Khi phụ tải vào cao điểm tối",
    ],
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "EVN Data");
  XLSX.writeFile(workbook, "evn_data.xlsx");
}

async function main() {
  try {
    const data = await crawlPages();
    if (data.length > 0) {
      saveToExcel(data);
      console.log("Data saved to evn_data.xlsx");
    } else {
      console.log("No data collected.");
    }
  } catch (error) {
    console.error("Error in main process:", error.message);
  }
}

main();
