const message = document.querySelector('#preprep-message');
const userInput = document.querySelector('#user-name');

function getUser() {
  const saved = localStorage.getItem('queueUser') || '';
  if (userInput && !userInput.value) userInput.value = saved;
  const user = userInput ? userInput.value.trim() : saved;
  if (!user) throw new Error('Enter your name or initials first.');
  localStorage.setItem('queueUser', user);
  return user;
}

async function postJson(url, body) {
  const response = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  const data = await response.json();
  if (!response.ok || data.ok === false) throw new Error(data.message || 'Request failed.');
  return data;
}

function renderQueues(queues) {
  for (const queueName of ['prep', 'qa']) {
    const list = document.querySelector(`#${queueName}-queue`);
    if (!list) continue;
    list.innerHTML = '';
    for (const entry of queues[queueName] || []) {
      const li = document.createElement('li');
      li.textContent = entry.user;
      list.appendChild(li);
    }
  }
}

if (document.querySelector('#preprep-form')) {
  document.querySelector('#preprep-form').addEventListener('submit', async (event) => {
    event.preventDefault();
    const form = event.currentTarget;
    const body = Object.fromEntries(new FormData(form).entries());
    body.ready_for_prep = form.ready_for_prep.checked;
    try {
      const data = await postJson('/api/preprep/rows', body);
      message.textContent = `Added row ${data.id} to the Excel queue.`;
      setTimeout(() => location.reload(), 900);
    } catch (error) {
      message.textContent = error.message;
    }
  });
}

if (userInput) {
  userInput.value = localStorage.getItem('queueUser') || '';
}

document.addEventListener('click', async (event) => {
  const button = event.target.closest('button');
  if (!button) return;
  try {
    if (button.dataset.join) {
      const data = await postJson(`/api/queue/${button.dataset.join}/join`, { user: getUser() });
      renderQueues(data.queues);
      return;
    }
    const card = button.closest('.order-card');
    if (!card) return;
    const row = card.dataset.row;
    const action = button.dataset.assign ? 'assign-next' : button.dataset.claim ? 'claim' : button.dataset.complete ? 'complete' : button.dataset.skip ? 'skip' : null;
    const queue = button.dataset.assign || button.dataset.claim || button.dataset.complete || button.dataset.skip;
    if (!action || !queue) return;
    const body = action === 'assign-next' ? { queue } : { queue, user: getUser() };
    const data = await postJson(`/api/order/${row}/${action}`, body);
    if (data.queues) renderQueues(data.queues);
    alert(data.message || 'Saved.');
    location.reload();
  } catch (error) {
    alert(error.message);
  }
});
