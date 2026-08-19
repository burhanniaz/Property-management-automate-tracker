require("dotenv").config();
const express = require("express");
const dashboardRouter = require("./routes/dashboard");

const app = express();

app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

app.use("/api/dashboard", dashboardRouter);

app.use((req, res) => {
  res.status(404).json({ error: "Not found" });
});

// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: err.message });
});

const port = Number(process.env.PORT || 3000);
app.listen(port, () => {
  console.log(`property-tracker-backend listening on port ${port}`);
});
