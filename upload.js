const express = require('express');
const router = express.Router();
const multer = require('multer');
const xlsxParse = require('./xlsxParse');

const storage = multer.diskStorage({
  destination: './public/uploads/',
  filename: function(req, file, cb) {
    cb(null, 'data.xlsx');
  }
});

const upload = multer({
  storage
}).single('excel');

router.post('/', (req, res) => {
  upload(req, res, err => {
    if (err) {
      res.json({
        msg: err
      });
    } else {
      if (req.file === undefined) {
        res.json({
          msg: 'Error: No File Selected!'
        });
      } else {
        const file = xlsxParse();
        res.json(file);
      }
    }
  });
});

module.exports = router;
