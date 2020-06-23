var express = require('express');
var fs = require("fs");
var request = require('request');
var bodyParser = require('body-parser');
var json2xls = require('json2xls');
var datetime = require('node-datetime');
var app = express();

app.use(bodyParser.urlencoded({ "extended": false }));
app.use(bodyParser.json());

const FILENAME = "all_members_";
const STORAGE = "/excels";

var role = 0;
var url = "https://api.scrapcatapp.com/kpa/get-all-members/export-to-excel";

app.use(json2xls.middleware);

app.get('/', (req, res) => {
    res.sendfile('index.html');
});

app.get('/all-users-export-to-excel-download', (req, res) => {
    get_all_members(res, true, role);
    res.setHeader('Content-Type', 'application/json');
});

app.get('/subscribe-unsubscribe-list-generate', (req, res) => {
    get_subscribe_unsubscribe_list(res);
    res.setHeader('Content-Type', 'application/json');
});

app.get('/check-point', (req, res) => {

    var isValid = false;
    var dropbox_link = "https://www.dropbox.com/sh/0jx9qa3hpj0yepr/AACCapL-fHbjsE_4SX9L-Jkfa?dl=0";

    // super admin
    if (req.query["password"] == "$65q9iV20ZJTTA0HQlQ9VBbVN9$gub") {
        isValid = true;
        role = 999;
        dropbox_link = "https://www.dropbox.com/sh/20aissro9ml3rha/AAAHeOTlghwhLYf-fSSig0pRa?dl=0";
    }

    // admin
    if (req.query["password"] == "o2scNdy9u@Qk4^KA") {
        isValid = true;
        role = 1;
    }

    res.setHeader('Content-Type', 'application/json');
    res.send(JSON.stringify({ status: isValid, role: role, link: dropbox_link }));
});

//nNfbfOTMzDFHCYw$K6ai72Yz0^SCxmXo#3ufs&^&e$QQiK6TMNr$2107SjMQ

function get_all_members(response, IsDownload, role) {
    let request_body = {
        "request_id": 1234567890,
        "session_id": "DS8SAD545489122er4"
    };

    console.log(role);

    if (role > 900) {
        url = "https://api.scrapcatapp.com/kpa/get-all-members/export-to-excel?show-password=true";
    }
    else if (role > 0 && role < 900) {
        url = "https://api.scrapcatapp.com/kpa/get-all-members/export-to-excel";
    }
    else {
        response.send({ status: 405, message: "You don't have any permissions to access this page, we got your location, and we are tracing you now. Got it? Good luck.."});
        return false;
    }

    request({
        "uri": url,
        "method": "GET",
        "json": request_body
    }, (err, res, body) => {
        if (!err) {
            var count = body.count;
            console.log('count: ' + count);
            if (count > 0) {
                var members = body.data;
                console.log('members: ' + members);
                save_into_excel(response, members, IsDownload, role);
                return false;
            }
            response.send("Oops, something went wrong.");
        } else {
            console.error("Unable to send message:" + err);
            response.send({ status: 403, message: "Unable to send message: " + err });
        }
    });
}

function get_subscribe_unsubscribe_list(response) {

    let request_body = {
        "request_id": 1234567890,
        "session_id": "DS8SAD545489122er4"
    };

    url = "https://sca-api-staging1.scrapcat.net/kpa/subscribe-unsubscribe-list/export-to-excel";

    request({
        "uri": url,
        "method": "GET",
        "json": request_body
    }, (err, res, body) => {

            console.log(body.data);

        if (!err) {
            var members = body.data;
            console.log('members: ' + members);
            save_into_excel2(response, members);
            return false;
        } else {
            console.error("Unable to send message:" + err);
            response.send({ status: 403, message: "Unable to send message: " + err });
        }
    });

}

function save_into_excel(response, members, IsDownload, role) {
    var dt = datetime.create();
    var formatted = dt.format('mdY_HMS');

    var _filename = FILENAME + formatted;

    var _role = role == 999 ? "private" : "public";

    var _storage = "./excels/" + _role + "/" + _filename;

    var xls = json2xls(members);
    fs.writeFileSync(_storage + '.xlsx', xls, 'binary');

    if (IsDownload) {
        // response.xls(_filename + '.xlsx', members);
        response.send({ status: 200, message: "The excel file was generated and ready to download, please click \"Excel Archives\" to see it." });
        return false;
    }
    response.send({ status: 200, message: "Export to excel has been done.\r\nFilename: " + _filename + ".xlsx"});
}

function save_into_excel2(response, members) {
    var dt = datetime.create();
    var formatted = dt.format('mdY_HMS');

    var _filename = FILENAME + formatted;
    var _role = "private/Subscribe-Unsubscribe";
    var _storage = "./excels/" + _role + "/" + _filename;

    var xls = json2xls(members);
    fs.writeFileSync(_storage + '.xlsx', xls, 'binary');

    response.send({ status: 200, message: "The excel file was generated and ready to download, please click \"Excel Archives\" to see it." });
    return false;
}

// app.listen(7878);
app.listen(process.env.PORT);