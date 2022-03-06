const express = require("express");
const app = express();
const axios = require("axios").default;
const ExcelJS = require("exceljs");

app.get("/", (req, res) => {
  const data = async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet", { rightToLeft: true });
    sheet.columns = [
      { header: "نام محصول", key: "name" },
      { header: "لینک محصول", key: "permalink" },
      { header: "تاریخ ایجاد محصول", key: "date_created" },
      { header: "وضعیت انتشار", key: "status" },
      { header: "کد محصول", key: "sku" },
      { header: "قیمت مشتری", key: "price" },
      { header: "تعداد فروش", key: "total_sales" },
      { header: "وضعیت موجودی", key: "stock_status" },
    ];
    let final = [];
    for (let i = 1; i <= 8; i++) {
      try {
        let products = await axios.get(
          `https://site.com/wp-json/wc/v3/products?per_page=100&page=${i}`,
          {
            headers: {
              Authorization: "Basic ",
            },
          }
        );
        products.data.forEach((product) => {
          final.push({
            name: product.name,
            permalink: product.permalink,
            date_created: product.date_created
              ? product.date_created
              : "نامشخص",
            status: product.status === "publish" ? "منتشر شده" : "پیش نویس",
            sku: product.sku,
            price: Intl.NumberFormat("en-US").format(
              parseInt(product.price) + "0"
            ),
            total_sales: parseInt(product.total_sales),
            stock_status:
              product.stock_status == "instock" ? "موجود" : "ناموجود",
          });
        });
      } catch (e) {
        console.log(e);
      }
    }
    sheet.addRows(final);
    sheet.getRow(1).font = { bold: true };
    await workbook.xlsx.writeFile("products.xlsx");
    res.json({
      length: final.length,
      data: final,
    });
  };
  data();
});

app.listen(3000, () => {
  console.log("app running");
});
