const env = require("./env");
const createError = require('http-errors');
const express = require('express');
const path = require('path');
const cookieParser = require('cookie-parser');
const logger = require('./middleware/logger');
const parser = require('./routes/parser');

const handleException = (error, _, res, __) => {
  logger.error(error.stack);
  // Any error which is not triggereed by 
  // application logic is considered as a errors internal to server side
  // This means if any code which does not handle application exceptions would be 
  // considered as internal error (which is wrong implementation of application exception logic)
  if(!error.applicationError) {
      error.message = 'Internal Server Error';
      error.statusCode = 500;
  }
  
  res.status(error.statusCode || 500).json({ error: error.message || error.toString() });
}

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/parser', parser, handleException);

// catch 404 and forward to error handler
app.use(function(req, res, next) {
  next(createError(404));
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  res.status(err.statusCode || 500).json({ error: err.message || err.toString() });
});

module.exports = app;