var express = require('express');
var multer = require('multer');
var app = express();
var Trans = require("./TransPlugin");

var lastUpload = "";

var storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/')
    },
    filename: function (req, file, cb) {
        var ext = file.originalname.split(".");
        var p = file.fieldname + '-' + Date.now() + "." + ext[ext.length-1];
        cb(null, p)
        lastUpload = p;
    }
});

var upload = multer({storage:storage });


app.get('/', function (req, res) {
    res.send("<h3>taoqi excel转换小工具</h3><form action='upload' method='post' enctype='multipart/form-data'><input type='file' name='file' /><br/><br/><input type='submit' value='上传'/></form>");
});

app.post('/upload', upload.single('file'), function(req, res, next) {
    


    res.download(Trans.TransExcel("uploads/"+lastUpload));
});

var server = app.listen(12345, function () {
  var host = server.address().address;
  var port = server.address().port;

  console.log('Example app listening at http://%s:%s', host, port);
});