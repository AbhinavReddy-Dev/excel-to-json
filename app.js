//requiring packages
const express = require("express");
const app = express();
const bodyParser = require("body-parser");
//middleware multer to save the file and process
const multer = require("multer");
const xlstojson = require("xls-to-json-lc");
const xlsxtojson = require("xlsx-to-json-lc");

app.use(bodyParser.json());

var storage = multer.diskStorage({
  //multer disk storage settings, cb is callback
  destination: function(req, file, cb) {
    cb(null, "./uploads/");
  },
  filename: function(req, file, cb) {
    var datetimestamp = Date.now();
    cb(
      null,
      file.fieldname +
        "-" +
        datetimestamp +
        "." +
        file.originalname.split(".")[file.originalname.split(".").length - 1]
    );
  }
});

var upload = multer({
  //multer settings
  storage: storage,
  fileFilter: function(req, file, callback) {
    //file filter
    if (
      ["xls", "xlsx"].indexOf(
        file.originalname.split(".")[file.originalname.split(".").length - 1]
      ) === -1
    ) {
      return callback(new Error("Wrong extension type"));
    }
    callback(null, true);
  }
}).single("file"); //enables only one file to be uploaded

app.get("/", function(req, res) {
  res.sendFile(__dirname + "/index.html");
});

/** API path that will upload the files */
app.post("/upload", function(req, res) {
  var exceltojson;
  upload(req, res, function(err) {
    if (err) {
      res.json({ error_code: 1, err_desc: err });
      return;
    }
    /** Multer gives us file info in req.file object */
    if (!req.file) {
      res.json({ error_code: 1, err_desc: "No file passed" });
      return;
    }
    /** Check the extension of the incoming file */
    if (
      req.file.originalname.split(".")[
        req.file.originalname.split(".").length - 1
      ] === "xlsx"
    ) {
      exceltojson = xlsxtojson;
    } else {
      exceltojson = xlstojson;
    }
    exceltojson(
      {
        input: req.file.path,
        output: "output.json", //null else  "output.json"
        lowerCaseHeaders: true
      },
      function(err, result) {
        if (err) {
          console.error(err);
        } else {
          console.log(result);
          res.json(result);
          // return res.json(result);
        }
      }
    );
  });
  app.redirect("/index.html");
});

app.listen("3000", function() {
  console.log("running on 3000...");
});
