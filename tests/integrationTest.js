var assert = require('assert');
var https = require('https');
var querystring = require('querystring');
var graphHelper = require('../utils/graphHelper.js');
var emailer = require('../utils/emailer.js');
var path = require('path');
var fs = require('fs');

describe('Integration', function () { // eslint-disable-line no-undef
  var accessToken;
  before( // eslint-disable-line no-undef
    function (done) {
      // Read variables from testConfig.json file
      var configFilePath = path.join(__dirname, 'testConfig.json');
      var config = JSON.parse(fs.readFileSync(configFilePath, { encoding: 'utf8' }));

      var postData = querystring.stringify(
        {
          grant_type: 'password',
          resource: 'https://graph.microsoft.com/',
          client_id: '15e94743-f2e0-4dcd-9695-4901a0175fc6',
          client_secret: '459rqch7iCCkaaO2gF7jbRb',
          username: 'landry.kammogne@improving.com',
          password: 'K@lamar2016'
        }
      );

      var postOptions = {
        host: 'login.microsoftonline.com',
        port: '443',
        path: '/common/oauth2/token',
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Content-Length': Buffer.byteLength(postData)
        }
      };

      var postRequest = https.request(postOptions, function (res) {
        var data = '';
        res.setEncoding('utf8');
        res.on('data', function (chunk) {
          data += chunk;
        });
        res.on('end', function () {
          accessToken = JSON.parse(data).access_token;
          done();
        });
      });

      postRequest.on('error', function (e) {
        console.log('Error: ' + e.message);
        done(e);
      });

      postRequest.write(postData);
      postRequest.end();
    }
  );
  it( // eslint-disable-line no-undef
    'Checking that the sample can send an email',
    function (done) {
        console.log('Test Landry');
        var picPath = path.join(__dirname, '/Mypic.jpg');
      var postBody = emailer.generateMailBody(
        'Landry Kammogne',
          'landry.kammogne@improving.com',
          'https://www.google.com/imgres?imgurl=http%3A%2F%2Fdallas-csharp-sig.com%2Fimg%2FImproving-logo-lg.jpg&imgrefurl=http%3A%2F%2Fdallas-csharp-sig.com%2F&docid=lM5AABxCY6yqIM&tbnid=ScqTq1XJT1DB-M%3A&vet=10ahUKEwi05-DzkonUAhUQ84MKHVKnC5wQMwglKAAwAA..i&w=2708&h=849&bih=885&biw=1745&q=improving%20enterprises&ved=0ahUKEwi05-DzkonUAhUQ84MKHVKnC5wQMwglKAAwAA&iact=mrc&uact=8',
          picPath
      );
      graphHelper.postSendMail(
        accessToken,
        JSON.stringify(postBody),
        function (error) {
          assert(error === null, '\nThe sample failed to send an email');
          done();
        });
    }
  );
});
