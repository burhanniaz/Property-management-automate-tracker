const express = require("express");
const { fetchOneKpi, getAllKpisCached } = require("../sheetsClient");

const router = express.Router();

router.get("/", async (req, res, next) => {
  try {
    const { data, cached, ageSeconds } = await getAllKpisCached();
    res.json({
      data,
      generatedAt: new Date().toISOString(),
      cached,
      cacheAgeSeconds: ageSeconds,
    });
  } catch (err) {
    next(err);
  }
});

router.get("/:name", async (req, res, next) => {
  try {
    const value = await fetchOneKpi(req.params.name);
    if (value === undefined) {
      return res.status(404).json({ error: `Unknown named range: ${req.params.name}` });
    }
    res.json({ name: req.params.name, value, generatedAt: new Date().toISOString() });
  } catch (err) {
    next(err);
  }
});

module.exports = router;
