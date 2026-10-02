const API_URL = import.meta.env.VITE_API_URL || 'http://10.1.100.69:47831';

async function request(path, options = {}) {
  try {
    const response = await fetch(`${API_URL}${path}`, {
      ...options,
      headers: { 'Content-Type': 'application/json', ...(options.headers || {}) }
    });
    const payload = await response.json();
    if (!response.ok) return { data: null, error: { message: payload.error || response.statusText, status: response.status } };
    return { data: payload.data, error: null };
  } catch (error) {
    return { data: null, error };
  }
}

class Query {
  constructor(table) { this.table = table; this.action = 'select'; this.body = null; this.filters = []; this.orders = []; this.head = false; }
  select(columns = '*', options = {}) { this.action = 'select'; this.columns = columns; this.head = Boolean(options.head); return this; }
  insert(value) { this.action = 'insert'; this.body = value; return this; }
  update(value) { this.action = 'update'; this.body = value; return this; }
  delete() { this.action = 'delete'; return this; }
  eq(column, value) { this.filters.push({ op: 'eq', column, value }); return this; }
  in(column, value) { this.filters.push({ op: 'in', column, value }); return this; }
  not(column, op, value) { this.filters.push({ op: 'not', column, value, operand: op }); return this; }
  order(column, options = {}) { this.orders.push({ column, ascending: options.ascending !== false }); return this; }
  limit(value) { this.limitValue = value; return this; }
  then(resolve, reject) {
    return request(`/api/tables/${this.table}`, {
      method: this.action === 'select' ? 'POST' : this.action === 'insert' ? 'PUT' : this.action === 'update' ? 'PATCH' : 'DELETE',
      body: JSON.stringify({ rows: this.body, filters: this.filters, orders: this.orders, limit: this.limitValue, head: this.head })
    }).then(resolve, reject);
  }
}

class Channel {
  constructor() { this.handlers = []; this.timer = null; }
  on(type, _filter, callback) { if (type === 'postgres_changes') this.handlers.push(callback); return this; }
  subscribe(callback) {
    callback?.('SUBSCRIBED');
    if (this.handlers.length) this.timer = setInterval(() => this.handlers[0]({ eventType: 'UPDATE' }), 5000);
    return this;
  }
  presenceState() { return {}; }
  async track() { return 'ok'; }
  async untrack() { return 'ok'; }
}

export function createApiClient() {
  return {
    from: (table) => new Query(table),
    rpc: (name, body = {}) => request(`/api/rpc/${name}`, { method: 'POST', body: JSON.stringify(body) }),
    channel: () => new Channel(),
    removeChannel: (channel) => { if (channel?.timer) clearInterval(channel.timer); }
  };
}

export { API_URL };
