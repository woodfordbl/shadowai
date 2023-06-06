const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

// Define routes
app.get('/', (req, res) => {
  res.sendFile(__dirname + 'client/assets/icon-128.png');
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
