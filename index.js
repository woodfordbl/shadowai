const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

// Serve the static files
app.use('/client', express.static('client'));
app.use('/client', express.static('client/src/taskpane'));
app.use('/client', express.static('client/assets'));


// Define routes
app.get('/', (req, res) => {
  res.send("Hello World!");
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
