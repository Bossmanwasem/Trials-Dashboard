import { createServer } from './api.js';
const port = Number(process.env.PORT || 47831);
createServer().listen(port, '10.1.100.69', () => console.log(`Trials Dashboard API listening at http://10.1.100.69:${port}`));
