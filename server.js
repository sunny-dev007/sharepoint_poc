const express = require('express');
const session = require('express-session');
const path = require('path');
const bodyParser = require('body-parser');
const cors = require('cors');
require('dotenv').config();

global.appRoot = path.resolve(__dirname);

const app = express();

//Logo
app.use('/images', express.static(__dirname + '/images'));

/**
 * Normalize a port into a number, string, or false.
 */

function normalizePort(val) {
  var port = parseInt(val, 10);

  if (isNaN(port)) {
    // named pipe
    return val;
  }

  if (port >= 0) {
    // port number
    return port;
  }

  return false;
}

const PORT = normalizePort(process.env.PORT || '8000');
//Need to add the following


if (!(process.env["Prod"])){
  //In Dev
  process.env.appUrl = 'http://localhost:8080';
  console.log("Dev Server Started");
}

const startServer = () => {

      // middleware function to check for logged-in users
      var sessionChecker = (req, res, next) => {
        next();
      };

      app.use(cors());

      app.use(bodyParser.json({
        limit: '100mb'
      }));

      // Importing the Hive 2.0 APIs Routing...
      require("./api/v1/routes/sharepoint")(app);
     
      app.use('/', sessionChecker, express.static(path.join(__dirname,'client/build')));
      app.use('/static', sessionChecker, express.static(path.join(__dirname,'client/build/static')));

      app.get('/*', function(req, res) {
        res.sendFile(path.join(__dirname, 'client/build/index.html'), function(err) {
          if (err) {
            res.status(500).send(err)
          }
        })
      })

      //app.listen();
      //Temp Solution for Linux vNet process.env.PORT should be PORT constant
      var server = require('http').createServer(app)

        server.listen(PORT, () => {
          console.log(`app listening on port ${PORT}!`);
        });
}



startServer()
