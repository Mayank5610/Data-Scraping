const puppeteer = require("puppeteer");
const xl = require("excel4node");
const wb = new xl.Workbook();
const ws = wb.addWorksheet("MRERA");

//https://maharerait.mahaonline.gov.in/SearchList/Search

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
  });
  const page = await browser.newPage();
  await page.goto(
    "https://maharerait.mahaonline.gov.in/searchlist/search?MenuID=1069",
    {
      waitUntil: "networkidle2",
      timeout: 0,
    }
  );

  await page.waitForSelector("#Promoter");
  await page.click("#Promoter");

  await page.waitForSelector("#btnAdvance");
  await page.click("#btnAdvance");

  await page.waitForSelector("#State");
  await page.select("#State", "27");

  await page.waitForSelector("#District");
  await page.select("#District", "519");

  await page.waitForSelector("#PType");
  await page.select("#PType", "13");

  await page.waitForSelector("#btnSearch");
  await page.click("#btnSearch");

  await page.waitForSelector("table");
  const scrapedData = [];

  console.log("First Time");
  try {
    await page.waitForSelector(".col-md-3.col-sm-3.text-center");
    const data = await page.evaluate(() => {
      const list = document.querySelector(".col-md-3.col-sm-3.text-center");
      const lists = list.innerText;

      const array = lists.match(/[0-9]+/g);
      return array;
    });
    // console.log(data[1]);
    let count = 1;

    while (count <= data[1]) {
      console.log("count: ", count);
      await page.waitForSelector("table");

      const rows = await page.$$("table tbody tr");

      const extractTableCellText = async (element) => {
        return await element.evaluate((node) => node.innerText);
      };

      for (let i = 0; i < rows.length; i++) {
        const columns = await rows[i].$$("td");

        const ProjectName = await extractTableCellText(columns[1]);
        const PromoterName = await extractTableCellText(columns[2]);

        const link = await columns[4].$("a");

        const newPagePromise = new Promise((x) =>
          browser.once("targetcreated", (target) => x(target.page()))
        );
        // console.log("newPagePromise", newPagePromise);
        await link.click();

        const newPage = await newPagePromise;

        await newPage.waitForSelector(".col-md-3.col-sm-3"); // Wait for new tab to load

        const newData = await newPage.evaluate(() => {
          const lists = document.querySelectorAll(".col-md-3.col-sm-3");
          const listArr = [];

          lists.forEach((list) => {
            // console.log("lists");
            listArr.push(list.innerText);
          });

          //   const number = listArr[33];
          //   const website = listArr[35];
          //   const projStatus = listArr[83];

          return listArr; //[(number, website, projStatus)];
        });
        await newPage.close();

        //   console.log(newData);
        let number = ""; // = newData[0];
        let website = ""; //= newData[1];
        let porjectStatus = ""; //= newData[2];
        const status = ["Office Number", "Website URL", "Project Status"];
        let getValue = false;
        //   let define = "";
        let index = 0;
        newData.forEach((item) => {
          if (getValue) {
            getValue = false;
            if (index === 1) {
              website = item;
            } else if (index === 2) {
              porjectStatus = item;
            } else {
              number = item;
            }
          }
          if (status.includes(item)) {
            if (item === "Office Number") {
              index = 0;
            } else if (item === "Website URL") {
              index = 1;
            } else if (item === "Project Status") {
              index = 2;
            }
            getValue = true;
          }
        });

        scrapedData.push({
          ProjectName,
          PromoterName,
          number,
          website,
          porjectStatus,
        });
        // console.log(scrapedData);
      }

      await page.waitForSelector("#btnNext");
      await page.evaluate(() => {
        // const btn removed because not used
        const nextButton = document.querySelector("#btnNext");
        if (nextButton) {
          const val =
            nextButton.hasAttribute("disabled") ||
            nextButton.classList.contains("disabled");
          if (!val) {
            nextButton.click();
            // console.log("Button Clicked!", btn);
          }
        }
      });

      const headers = Object.keys(scrapedData[0]);
      headers.forEach((header, columnIndex) => {
        ws.cell(1, columnIndex + 1).string(header); // Write headers to the first row
      });

      scrapedData.forEach((obj, rowIndex) => {
        headers.forEach((header, columnIndex) => {
          const value = obj[header]; // Get the value based on the header
          ws.cell(rowIndex + 2, columnIndex + 1).string(String(value)); // Write data to the worksheet
        });
      });

      wb.write("MRERA.xlsx", (err, stats) => {
        if (err) {
          console.log("Error occured while writing Excel file: ", err);
        } else {
          console.log(`Excel file "MRERA.xlsx" has been created successfully`);
        }
      });
      count++;
    }

    await browser.close(); // not running
  } catch (error) {
    console.log("Error: ", error);
  } finally {
    await browser.close();
  }
})();
