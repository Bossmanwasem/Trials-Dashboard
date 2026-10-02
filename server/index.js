import { createServer } from './api.js';
const port = Number(process.env.PORT || 47831);
createServer().listen(port, '127.0.0.1', () => console.log(`Trials Dashboard API listening at http://127.0.0.1:${port}`));
