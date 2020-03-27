const express = require("express");
const logger = require("morgan");
const cookieParser = require("cookie-parser");
const bodyParser = require("body-parser");

const cors = require("cors");
const fs = require("fs");
var xlsx = require("xlsx");
const xlsxParse = require("./xlsxParse");

//UPLOAD ROUTE
const upload = require("./upload");

const app = express();
app.use(logger("dev"));
app.use(cors());
app.use(bodyParser.json());
app.use(
  bodyParser.urlencoded({
    extended: false
  })
);
app.use(cookieParser());
app.use(express.static("./build"));
app.use(express.static("./public"));

//app.use(express.static(path.join(__dirname, "build")));

app.get("/*", (req, res) => {
  res.sendFile(path.join(__dirname, "build", "index.html"));
});

app.get("/", (req, res) => {
  let file = xlsxParse();

  res.status(200).json(file);
});

//ENDPOINT FOR IMPORTING FILE
app.use("/upload", upload);

//ENDPOINT FOR EXPORTING FILE

app.post("/json", (req, res) => {
  var data;
  data = req.body;

  // data.map(element => {
  //   console.log(element.Requirement);
  //   return element;
  // });
  // console.log(data);

  var newWB = xlsx.utils.book_new();
  var newWS = xlsx.utils.json_to_sheet(data);
  xlsx.utils.book_append_sheet(newWB, newWS, "Company");

  console.log(newWB.SheetNames);
  xlsx.writeFile(newWB, "./public/downloads/newData.xlsx");
  const filePath = `${__dirname}/public/downloads/newData.xlsx`;
  console.log(filePath);

  res.download(filePath);
});

// app.post('/export', (req, res) => {
//   var newWB = xlsx.utils.book_new();
//   var newWS = xlsx.utils.json_to_sheet(file);
//   console.log(newWB.SheetNames);

//   xlsx.utils.book_append_sheet(newWB, newWS, 'Company');
//   console.log(newWB.SheetNames);
//   xlsx.writeFile(newWB, 'newData.xlsx');
//   res.status(200).send('Excel file generated successsfully');
// });

app.listen(8080, () => {
  console.log("Server is up and running on port", 8080);
});
