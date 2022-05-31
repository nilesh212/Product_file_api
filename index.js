const { urlencoded } = require("express");
const express = require("express");
const multer = require("multer");
const Excel = require("exceljs");
const fs = require("fs");
const axios = require("axios");
const port = 8000;
const app = express();

app.use(express.urlencoded({ extended: false }));

const upload = multer({ dest: "./Product_file" });

app.post("/", upload.single("product_file"), async function (req, res) {
  try {
    let file_extenstion = req.file.originalname.split(".")[1];

    // If the file is not excel sheet then return 409 error.
    if (
      file_extenstion != "xlsx" &&
      file_extenstion != "xls" &&
      file_extenstion != "xlsm" &&
      file_extenstion != "xlt"
    ) {
      fs.unlinkSync(`./Product_file/${req.file.filename}`);

      return res.status(409).json({
        data: {
          ErrorCode: 409,
          Error: "Please upload right excel file",
        },
      });
    }

    // Modifying File
    const workbook = new Excel.Workbook();

    await workbook.xlsx
      .readFile(`./Product_file/${req.file.filename}`)
      .then(async function () {
        let worksheet = workbook.getWorksheet(1);
        // Below loop goes through each row of file and assigns the value to price column based on product_name
        for (let index = 2; index <= worksheet._rows.length; index++) {
          let row = worksheet.getRow(index);
          let product_name = row.getCell(1).value;
          let url = `https://api.storerestapi.com/products/${product_name}`;
          await axios
            .get(url)
            .then((response) => {
              row.getCell(2).value = response.data.data.price;
              console.log("Row Updated");
            })
            .catch((error) => {
              console.log("error");
            });
          row.commit();
        }

        // write the data into the same file
        return workbook.xlsx.writeFile(`./Product_file/${req.file.filename}`);
      })
      .catch(function (err) {
        console.log("This is error: ", err);
      });

    //Sending to the user
    let fileName = `./Product_file/${req.file.filename}`;
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=" + fileName);
    await workbook.xlsx.write(res);

    // Removing file from local machine
    // If you want to store the file then comment line number 75
    fs.unlinkSync(`./Product_file/${req.file.filename}`);

    // return the response
    return res.end();
  } catch (error) {
    console.log("**********************ERROR: ", error);
    return res.status(500).json({
      data: {
        Error: "Internal Server Error",
      },
    });
  }
});

app.listen(port, function (err) {
  if (err) {
    console.log("Error at running the port: \n", err);
    return;
  }
  console.log("Server running at port: ", port);
  return;
});
