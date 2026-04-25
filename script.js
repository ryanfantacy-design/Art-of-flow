const DEFAULT_TASKS = [
  { id: "gym", name: "Gym", points: 10 },
  { id: "maths-stats", name: "Maths/Stats", points: 10 },
  { id: "commerce", name: "Commerce", points: 10 },
  { id: "reasoning", name: "Reasoning", points: 10 },
  { id: "polity-economics", name: "Polity/Economics", points: 10 },
  { id: "current-affairs", name: "Current Affairs", points: 10 },
  { id: "revision", name: "Revision", points: 10 },
  { id: "coaching", name: "Coaching", points: 10 },
  { id: "mains-study", name: "Mains Study (Banking/Auditing/BO)", points: 10 },
  { id: "answer-writing", name: "Answer Writing", points: 10 }
];

const PHASES = [
  { name: "Phase 1", range: "Day 1-20", title: "Foundation + Catch-up", min: 1, max: 20 },
  { name: "Phase 2", range: "Day 21-50", title: "Full Coverage + Practice", min: 21, max: 50 },
  { name: "Phase 3", range: "Day 51-70", title: "Mock Tests + Revision", min: 51, max: 70 }
];

const TOPICS = {
  "Commercial Maths": ["Percentage", "Ratio", "Mean", "Median", "SD"],
  Reasoning: ["Series", "Coding", "Direction", "Blood relation"],
  Commerce: ["Business Organisation", "Banking", "Auditing"],
  Polity: ["CAG", "Finance Commission", "Budget"],
  Economics: ["GDP", "Inflation", "Monetary Policy"]
};

const STORAGE_KEY = "fao-complete-study-tracker";
const REVIEW_WEEKS = 10;
const MOCK_TESTS = 3;

const todayLabel = document.getElementById("todayLabel");
const scoreValue = document.getElementById("scoreValue");
const scoreMax = document.getElementById("scoreMax");
const statusMessage = document.getElementById("statusMessage");
const streakValue = document.getElementById("streakValue");
const heroMomentum = document.getElementById("heroMomentum");
const taskList = document.getElementById("taskList");
const resetButton = document.getElementById("resetButton");
const editorList = document.getElementById("editorList");
const addTaskButton = document.getElementById("addTaskButton");
const weeklyTotal = document.getElementById("weeklyTotal");
const weeklyAverage = document.getElementById("weeklyAverage");
const weeklySummaryBody = document.getElementById("weeklySummaryBody");
const weeklyGraph = document.getElementById("weeklyGraph");
const currentPhase = document.getElementById("currentPhase");
const phaseGrid = document.getElementById("phaseGrid");
const weeklyReviewBody = document.getElementById("weeklyReviewBody");
const topicSummaryGrid = document.getElementById("topicSummaryGrid");
const topicGroups = document.getElementById("topicGroups");
const topicEditorList = document.getElementById("topicEditorList");
const addTopicButton = document.getElementById("addTopicButton");
const mockTestsBody = document.getElementById("mockTestsBody");
const addMockTestButton = document.getElementById("addMockTestButton");
const journalEntry = document.getElementById("journalEntry");

function getDateKey(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDate(date = new Date()) {
  return new Intl.DateTimeFormat("en-IN", {
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric"
  }).format(date);
}

function createTaskId() {
  return `task-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function createEmptyDay(tasks) {
  return tasks.reduce((acc, task) => {
    acc[task.id] = {
      done: false,
      startHour: "",
      startMinute: "",
      startMeridiem: "AM",
      endHour: "",
      endMinute: "",
      endMeridiem: "PM"
    };
    return acc;
  }, {});
}

function createDefaultTopics() {
  return Object.fromEntries(
    Object.entries(TOPICS).map(([subject, topics]) => [
      subject,
      Object.fromEntries(topics.map((topic) => [topic, false]))
    ])
  );
}

function createDefaultTopicItems() {
  return Object.entries(TOPICS).flatMap(([subject, topics]) =>
    topics.map((name) => ({
      id: createTaskId(),
      subject,
      name
    }))
  );
}

function createWeeklyReview() {
  return Array.from({ length: REVIEW_WEEKS }, (_, index) => ({
    week: index + 1,
    totalScore: "",
    avgScore: "",
    mockScore: "",
    weakArea: "",
    fixPlan: ""
  }));
}

function createMockTests() {
  return Array.from({ length: MOCK_TESTS }, (_, index) => ({
    testNo: index + 1,
    score: "",
    accuracy: "",
    weakArea: "",
    improvement: ""
  }));
}

function getMaxScore(tasks) {
  return tasks.reduce((total, task) => total + Number(task.points || 0), 0);
}

function getStatus(score) {
  if (score >= 80) {
    return { label: "Topper Mode", background: "#dcefe7", color: "#1c624b", className: "status-topper" };
  }

  if (score >= 60) {
    return { label: "Good", background: "#f6e7bf", color: "#85570d", className: "status-good" };
  }

  return { label: "Improve", background: "#f6ddd7", color: "#923d31", className: "status-improve" };
}

function defaultData() {
  return {
    appStartDate: getDateKey(),
    tasks: DEFAULT_TASKS,
    days: {},
    journal: {},
    weeklyReview: createWeeklyReview(),
    mockTests: createMockTests(),
    topics: createDefaultTopics(),
    topicItems: createDefaultTopicItems()
  };
}

function loadData() {
  const saved = localStorage.getItem(STORAGE_KEY);

  if (!saved) {
    return defaultData();
  }

  try {
    const parsed = JSON.parse(saved);
    const base = defaultData();
    return {
      ...base,
      ...parsed,
      tasks: Array.isArray(parsed.tasks) && parsed.tasks.length ? parsed.tasks : base.tasks,
      days: parsed.days || {},
      journal: parsed.journal || {},
      weeklyReview: Array.isArray(parsed.weeklyReview) && parsed.weeklyReview.length ? parsed.weeklyReview : base.weeklyReview,
      mockTests: Array.isArray(parsed.mockTests) && parsed.mockTests.length ? parsed.mockTests : base.mockTests,
      topics: parsed.topics || base.topics,
      topicItems: Array.isArray(parsed.topicItems) && parsed.topicItems.length ? parsed.topicItems : base.topicItems
    };
  } catch {
    return defaultData();
  }
}

function saveData(data) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
}

function migrateLegacyDay(day) {
  DEFAULT_TASKS.forEach((task) => {
    if (day[task.id] && typeof day[task.id] === "object") {
      return;
    }

    if (typeof day[task.id] === "boolean") {
      day[task.id] = {
        done: day[task.id],
        startHour: "",
        startMinute: "",
        startMeridiem: "AM",
        endHour: "",
        endMinute: "",
        endMeridiem: "PM"
      };
      return;
    }

    if (typeof day[task.name] === "boolean") {
      day[task.id] = {
        done: day[task.name],
        startHour: "",
        startMinute: "",
        startMeridiem: "AM",
        endHour: "",
        endMinute: "",
        endMeridiem: "PM"
      };
      delete day[task.name];
    }
  });
}

function syncDaysToTasks(data) {
  const validTaskIds = new Set(data.tasks.map((task) => task.id));

  Object.values(data.days).forEach((day) => {
    migrateLegacyDay(day);

    data.tasks.forEach((task) => {
      if (!day[task.id] || typeof day[task.id] !== "object") {
        day[task.id] = {
          done: false,
          startHour: "",
          startMinute: "",
          startMeridiem: "AM",
          endHour: "",
          endMinute: "",
          endMeridiem: "PM"
        };
      }

      if (typeof day[task.id].done !== "boolean") {
        day[task.id].done = false;
      }

      if (typeof day[task.id].startHour !== "string" && typeof day[task.id].startHour !== "number") {
        day[task.id].startHour = "";
      }

      if (typeof day[task.id].startMinute !== "string" && typeof day[task.id].startMinute !== "number") {
        day[task.id].startMinute = "";
      }

      if (typeof day[task.id].endHour !== "string" && typeof day[task.id].endHour !== "number") {
        day[task.id].endHour = "";
      }

      if (typeof day[task.id].endMinute !== "string" && typeof day[task.id].endMinute !== "number") {
        day[task.id].endMinute = "";
      }

      if (day[task.id].startMeridiem !== "AM" && day[task.id].startMeridiem !== "PM") {
        day[task.id].startMeridiem = "AM";
      }

      if (day[task.id].endMeridiem !== "AM" && day[task.id].endMeridiem !== "PM") {
        day[task.id].endMeridiem = "PM";
      }
    });

    Object.keys(day).forEach((taskId) => {
      if (!validTaskIds.has(taskId)) {
        delete day[taskId];
      }
    });
  });
}

function syncTopicData(data) {
  if (!Array.isArray(data.topicItems) || !data.topicItems.length) {
    data.topicItems = createDefaultTopicItems();
  }

  const topicState = {};

  data.topicItems.forEach((item) => {
    if (!topicState[item.subject]) {
      topicState[item.subject] = {};
    }

    const existingValue = data.topics?.[item.subject]?.[item.name];
    topicState[item.subject][item.name] = typeof existingValue === "boolean" ? existingValue : false;
  });

  data.topics = topicState;
}

function ensureToday(data) {
  const todayKey = getDateKey();

  if (!data.days[todayKey]) {
    data.days[todayKey] = createEmptyDay(data.tasks);
  }

  data.tasks.forEach((task) => {
    if (!data.days[todayKey][task.id] || typeof data.days[todayKey][task.id] !== "object") {
      data.days[todayKey][task.id] = {
        done: false,
        startHour: "",
        startMinute: "",
        startMeridiem: "AM",
        endHour: "",
        endMinute: "",
        endMeridiem: "PM"
      };
    }

    if (typeof data.days[todayKey][task.id].done !== "boolean") {
      data.days[todayKey][task.id].done = false;
    }

    if (typeof data.days[todayKey][task.id].startHour !== "string" && typeof data.days[todayKey][task.id].startHour !== "number") {
      data.days[todayKey][task.id].startHour = "";
    }

    if (typeof data.days[todayKey][task.id].startMinute !== "string" && typeof data.days[todayKey][task.id].startMinute !== "number") {
      data.days[todayKey][task.id].startMinute = "";
    }

    if (typeof data.days[todayKey][task.id].endHour !== "string" && typeof data.days[todayKey][task.id].endHour !== "number") {
      data.days[todayKey][task.id].endHour = "";
    }

    if (typeof data.days[todayKey][task.id].endMinute !== "string" && typeof data.days[todayKey][task.id].endMinute !== "number") {
      data.days[todayKey][task.id].endMinute = "";
    }

    if (data.days[todayKey][task.id].startMeridiem !== "AM" && data.days[todayKey][task.id].startMeridiem !== "PM") {
      data.days[todayKey][task.id].startMeridiem = "AM";
    }

    if (data.days[todayKey][task.id].endMeridiem !== "AM" && data.days[todayKey][task.id].endMeridiem !== "PM") {
      data.days[todayKey][task.id].endMeridiem = "PM";
    }
  });

  return todayKey;
}

function calculateScore(taskState, tasks) {
  return tasks.reduce((total, task) => total + (taskState[task.id]?.done ? Number(task.points || 0) : 0), 0);
}

function calculateStreak(days, tasks) {
  const maxScore = getMaxScore(tasks);
  let streak = 0;
  const cursor = new Date();

  while (true) {
    const key = getDateKey(cursor);
    const day = days[key];

    if (!day || calculateScore(day, tasks) !== maxScore) {
      break;
    }

    streak += 1;
    cursor.setDate(cursor.getDate() - 1);
  }

  return streak;
}

function calculateStudyDay(data) {
  const start = new Date(`${data.appStartDate}T00:00:00`);
  const today = new Date();
  const diff = Math.floor((today - start) / (1000 * 60 * 60 * 24));
  return Math.max(1, diff + 1);
}

function getCurrentPhase(day) {
  return PHASES.find((phase) => day >= phase.min && day <= phase.max) || {
    name: "Beyond Phase 3",
    range: `Day ${day}`,
    title: "Revision + Mock Focus"
  };
}

function renderTasks(tasks, taskState, onToggle, onTimeChange) {
  taskList.innerHTML = "";

  tasks.forEach((task) => {
    const item = document.createElement("label");
    item.className = "task-item";
    item.innerHTML = `
      <span class="task-main">
        <input type="checkbox" ${taskState[task.id]?.done ? "checked" : ""}>
        <span class="task-name">${task.name}</span>
      </span>
      <span class="task-right">
        <span class="time-cluster">
          <span class="time-cluster-label">Start</span>
          <span class="time-group">
            <input type="number" min="1" max="12" step="1" value="${taskState[task.id]?.startHour ?? ""}" data-task-time="startHour" data-id="${task.id}" placeholder="7">
            <span class="time-label">hr</span>
          </span>
          <span class="time-group">
            <input type="number" min="0" max="59" step="1" value="${taskState[task.id]?.startMinute ?? ""}" data-task-time="startMinute" data-id="${task.id}" placeholder="00">
            <span class="time-label">min</span>
          </span>
          <span class="time-group">
            <select data-task-time="startMeridiem" data-id="${task.id}">
              <option value="AM" ${taskState[task.id]?.startMeridiem === "AM" ? "selected" : ""}>AM</option>
              <option value="PM" ${taskState[task.id]?.startMeridiem === "PM" ? "selected" : ""}>PM</option>
            </select>
          </span>
        </span>
        <span class="time-cluster">
          <span class="time-cluster-label">End</span>
          <span class="time-group">
            <input type="number" min="1" max="12" step="1" value="${taskState[task.id]?.endHour ?? ""}" data-task-time="endHour" data-id="${task.id}" placeholder="8">
            <span class="time-label">hr</span>
          </span>
          <span class="time-group">
            <input type="number" min="0" max="59" step="1" value="${taskState[task.id]?.endMinute ?? ""}" data-task-time="endMinute" data-id="${task.id}" placeholder="00">
            <span class="time-label">min</span>
          </span>
          <span class="time-group">
            <select data-task-time="endMeridiem" data-id="${task.id}">
              <option value="AM" ${taskState[task.id]?.endMeridiem === "AM" ? "selected" : ""}>AM</option>
              <option value="PM" ${taskState[task.id]?.endMeridiem === "PM" ? "selected" : ""}>PM</option>
            </select>
          </span>
        </span>
        <span class="task-points">${task.points} points</span>
      </span>
    `;

    item.querySelector("input").addEventListener("change", (event) => onToggle(task.id, event.target.checked));
    item.querySelectorAll("[data-task-time]").forEach((input) => {
      input.addEventListener("input", (event) => onTimeChange(task.id, event.target.dataset.taskTime, event.target.value));
      input.addEventListener("change", (event) => onTimeChange(task.id, event.target.dataset.taskTime, event.target.value));
    });
    taskList.append(item);
  });
}

function renderTaskEditor(tasks) {
  editorList.innerHTML = "";

  tasks.forEach((task) => {
    const item = document.createElement("div");
    item.className = "editor-item";
    item.innerHTML = `
      <input type="text" value="${task.name}" data-action="edit-name" data-id="${task.id}" aria-label="Task name">
      <input type="number" min="0" step="1" value="${task.points}" data-action="edit-points" data-id="${task.id}" aria-label="Task points">
      <div class="editor-actions">
        <button type="button" class="delete-button" data-action="delete-task" data-id="${task.id}">Delete</button>
      </div>
    `;
    editorList.append(item);
  });
}

function renderWeeklySummary(days, tasks) {
  weeklySummaryBody.innerHTML = "";
  weeklyGraph.innerHTML = "";
  let total = 0;
  const maxScore = getMaxScore(tasks);

  for (let offset = 6; offset >= 0; offset -= 1) {
    const date = new Date();
    date.setDate(date.getDate() - offset);

    const key = getDateKey(date);
    const taskState = days[key] || {};
    const score = calculateScore(taskState, tasks);
    const completedCount = tasks.filter((task) => Boolean(taskState[task.id]?.done)).length;
    const status = getStatus(score);
    total += score;

    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${formatDate(date)}</td>
      <td>${completedCount} / ${tasks.length}</td>
      <td>${score} / ${maxScore}</td>
      <td><span class="status-chip ${status.className}">${status.label}</span></td>
    `;
    weeklySummaryBody.append(row);

    const barItem = document.createElement("div");
    barItem.className = "graph-bar-item";
    const barHeight = maxScore === 0 ? 10 : Math.max(10, Math.round((score / maxScore) * 140));
    barItem.innerHTML = `
      <span class="graph-score">${score}</span>
      <div class="graph-bar-wrap">
        <div class="graph-bar" style="height:${barHeight}px"></div>
      </div>
      <span class="graph-label">${formatDate(date).split(",")[0]}</span>
    `;
    weeklyGraph.append(barItem);
  }

  weeklyTotal.textContent = String(total);
  weeklyAverage.textContent = String(Math.round(total / 7));
}

function renderPhases(day) {
  phaseGrid.innerHTML = "";
  const activePhase = getCurrentPhase(day);
  currentPhase.textContent = `${activePhase.name}: ${activePhase.title}`;

  PHASES.forEach((phase) => {
    const card = document.createElement("article");
    card.className = `phase-card${activePhase.name === phase.name ? " active" : ""}`;
    card.innerHTML = `
      <p class="label">${phase.name}</p>
      <span class="phase-tag">${phase.range}</span>
      <h2>${phase.title}</h2>
      <p class="phase-copy">${activePhase.name === phase.name ? "Current focus phase." : "Planned study block."}</p>
    `;
    phaseGrid.append(card);
  });
}

function renderWeeklyReview(data) {
  weeklyReviewBody.innerHTML = "";

  data.weeklyReview.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.week}</td>
      <td><input class="table-input" data-group="weeklyReview" data-index="${index}" data-field="totalScore" value="${row.totalScore}"></td>
      <td><input class="table-input" data-group="weeklyReview" data-index="${index}" data-field="avgScore" value="${row.avgScore}"></td>
      <td><input class="table-input" data-group="weeklyReview" data-index="${index}" data-field="mockScore" value="${row.mockScore}"></td>
      <td><input class="table-input" data-group="weeklyReview" data-index="${index}" data-field="weakArea" value="${row.weakArea}"></td>
      <td><textarea class="table-textarea" data-group="weeklyReview" data-index="${index}" data-field="fixPlan">${row.fixPlan}</textarea></td>
    `;
    weeklyReviewBody.append(tr);
  });
}

function renderTopicSummary(topics) {
  topicSummaryGrid.innerHTML = "";

  Object.entries(topics).forEach(([subject, topicMap]) => {
    const list = Object.keys(topicMap);
    const done = list.filter((topic) => topicMap[topic]).length;
    const item = document.createElement("div");
    item.className = "mini-stat";
    item.innerHTML = `
      <span class="label">${subject}</span>
      <strong>${done} / ${list.length}</strong>
    `;
    topicSummaryGrid.append(item);
  });
}

function renderTopics(data) {
  topicGroups.innerHTML = "";
  renderTopicSummary(data.topics);

  Object.entries(data.topics).forEach(([subject, topicMap]) => {
    const list = Object.keys(topicMap);
    const done = list.filter((topic) => topicMap[topic]).length;
    const card = document.createElement("article");
    card.className = "topic-card";
    const listWrap = document.createElement("div");
    listWrap.className = "topic-list";

    list.forEach((topic) => {
      const checked = Boolean(data.topics[subject]?.[topic]);
      const row = document.createElement("label");
      row.className = `topic-row${checked ? " done" : ""}`;
      row.innerHTML = `
        <span class="topic-left">
          <input type="checkbox" ${checked ? "checked" : ""} data-subject="${subject}" data-topic="${topic}">
          <span class="topic-name">${topic}</span>
        </span>
        <span class="topic-mark">${checked ? "Done" : "Pending"}</span>
      `;
      listWrap.append(row);
    });

    card.innerHTML = `
      <div class="topic-header">
        <div>
          <p class="label">${subject}</p>
          <h2>${done} / ${list.length}</h2>
        </div>
        <span class="topic-progress">${Math.round((done / list.length) * 100)}%</span>
      </div>
    `;
    card.append(listWrap);
    topicGroups.append(card);
  });
}

function renderTopicEditor(data) {
  topicEditorList.innerHTML = "";

  data.topicItems.forEach((item) => {
    const row = document.createElement("div");
    row.className = "topic-editor-item";
    row.innerHTML = `
      <input type="text" value="${item.subject}" data-topic-action="edit-subject" data-id="${item.id}" aria-label="Subject name" placeholder="Subject name">
      <input type="text" value="${item.name}" data-topic-action="edit-name" data-id="${item.id}" aria-label="Topic name" placeholder="Topic name">
      <div class="editor-actions">
        <button type="button" class="delete-button" data-topic-action="delete-topic" data-id="${item.id}">Delete</button>
      </div>
    `;
    topicEditorList.append(row);
  });
}

function renderMockTests(data) {
  mockTestsBody.innerHTML = "";

  data.mockTests.forEach((test, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${test.testNo}</td>
      <td><input class="table-input" data-group="mockTests" data-index="${index}" data-field="score" value="${test.score}"></td>
      <td><input class="table-input" data-group="mockTests" data-index="${index}" data-field="accuracy" value="${test.accuracy}"></td>
      <td><input class="table-input" data-group="mockTests" data-index="${index}" data-field="weakArea" value="${test.weakArea}"></td>
      <td><textarea class="table-textarea" data-group="mockTests" data-index="${index}" data-field="improvement">${test.improvement}</textarea></td>
      <td><button type="button" class="delete-button" data-mock-action="delete-test" data-index="${index}">Delete</button></td>
    `;
    mockTestsBody.append(tr);
  });
}

function renderJournal(data, todayKey) {
  journalEntry.value = data.journal[todayKey] || "";
}

function updateUI(data, todayKey) {
  const todayTasks = data.days[todayKey];
  const maxScore = getMaxScore(data.tasks);
  const score = calculateScore(todayTasks, data.tasks);
  const status = getStatus(score);
  const studyDay = calculateStudyDay(data);

  scoreValue.textContent = String(score);
  scoreMax.textContent = String(maxScore);
  streakValue.textContent = String(calculateStreak(data.days, data.tasks));
  statusMessage.textContent = status.label;
  statusMessage.style.background = status.background;
  statusMessage.style.color = status.color;
  heroMomentum.textContent =
    score >= 80 ? "Topper Mode" :
    score >= 60 ? "Solid Progress" :
    score > 0 ? "Build Momentum" :
    "Start Strong";

  renderTasks(data.tasks, todayTasks, (taskId, checked) => {
    data.days[todayKey][taskId].done = checked;
    saveData(data);
    updateUI(data, todayKey);
  }, (taskId, field, value) => {
    if (field === "startMinute" || field === "endMinute") {
      data.days[todayKey][taskId][field] = value === "" ? "" : String(Math.min(59, Math.max(0, Number(value) || 0)));
    } else if (field === "startHour" || field === "endHour") {
      data.days[todayKey][taskId][field] = value === "" ? "" : String(Math.min(12, Math.max(1, Number(value) || 1)));
    } else if (field === "startMeridiem" || field === "endMeridiem") {
      data.days[todayKey][taskId][field] = value === "PM" ? "PM" : "AM";
    }
    saveData(data);
  });

  renderTaskEditor(data.tasks);
  renderWeeklySummary(data.days, data.tasks);
  renderPhases(studyDay);
  renderWeeklyReview(data);
  renderTopics(data);
  renderTopicEditor(data);
  renderMockTests(data);
  renderJournal(data, todayKey);
}

function resetDay(data, todayKey) {
  data.days[todayKey] = createEmptyDay(data.tasks);
  saveData(data);
  updateUI(data, todayKey);
}

function addTask(data, todayKey) {
  data.tasks.push({ id: createTaskId(), name: "New Task", points: 10 });
  syncDaysToTasks(data);
  ensureToday(data);
  saveData(data);
  updateUI(data, todayKey);
}

function handleTaskEditorInput(event, data, todayKey) {
  const target = event.target;
  const task = data.tasks.find((item) => item.id === target.dataset.id);

  if (!task) {
    return;
  }

  if (target.dataset.action === "edit-name") {
    task.name = target.value || "Untitled Task";
  }

  if (target.dataset.action === "edit-points") {
    task.points = Math.max(0, Number(target.value) || 0);
  }

  saveData(data);
}

function handleTaskEditorClick(event, data, todayKey) {
  const target = event.target;

  if (target.dataset.action !== "delete-task") {
    return;
  }

  data.tasks = data.tasks.filter((task) => task.id !== target.dataset.id);
  syncDaysToTasks(data);
  saveData(data);
  updateUI(data, todayKey);
}

function handleTableInput(event, data, todayKey) {
  const target = event.target;
  const group = target.dataset.group;
  const index = Number(target.dataset.index);
  const field = target.dataset.field;

  if (!group || Number.isNaN(index) || !field) {
    return;
  }

  if (group === "weeklyReview") {
    data.weeklyReview[index][field] = target.value;
  }

  if (group === "mockTests") {
    data.mockTests[index][field] = target.value;
  }

  saveData(data);
}

function handleTopicChange(event, data, todayKey) {
  const target = event.target;
  const subject = target.dataset.subject;
  const topic = target.dataset.topic;

  if (!subject || !topic) {
    return;
  }

  data.topics[subject][topic] = target.checked;
  saveData(data);
  updateUI(data, todayKey);
}

function addMockTest(data, todayKey) {
  data.mockTests.push({
    testNo: data.mockTests.length + 1,
    score: "",
    accuracy: "",
    weakArea: "",
    improvement: ""
  });
  saveData(data);
  updateUI(data, todayKey);
}

function handleMockTestClick(event, data, todayKey) {
  const target = event.target;

  if (target.dataset.mockAction !== "delete-test") {
    return;
  }

  const index = Number(target.dataset.index);

  if (Number.isNaN(index)) {
    return;
  }

  data.mockTests.splice(index, 1);
  data.mockTests = data.mockTests.map((test, testIndex) => ({
    ...test,
    testNo: testIndex + 1
  }));

  saveData(data);
  updateUI(data, todayKey);
}

function handleJournalInput(data, todayKey) {
  data.journal[todayKey] = journalEntry.value;
  saveData(data);
}

function handleTopicEditorInput(event, data, todayKey) {
  const target = event.target;
  const action = target.dataset.topicAction;
  const item = data.topicItems.find((topic) => topic.id === target.dataset.id);

  if (!action || !item) {
    return;
  }

  if (action === "edit-subject") {
    item.subject = target.value.trim() || "New Subject";
  }

  if (action === "edit-name") {
    item.name = target.value.trim() || "New Topic";
  }

  syncTopicData(data);
  saveData(data);
}

function refreshAfterEdit(data, todayKey) {
  syncDaysToTasks(data);
  syncTopicData(data);
  saveData(data);
  updateUI(data, todayKey);
}

function handleTopicEditorClick(event, data, todayKey) {
  const target = event.target;

  if (target.dataset.topicAction !== "delete-topic") {
    return;
  }

  data.topicItems = data.topicItems.filter((item) => item.id !== target.dataset.id);
  syncTopicData(data);
  saveData(data);
  updateUI(data, todayKey);
}

function addTopic(data, todayKey) {
  data.topicItems.push({
    id: createTaskId(),
    subject: "New Subject",
    name: "New Topic"
  });
  syncTopicData(data);
  saveData(data);
  updateUI(data, todayKey);
}

function init() {
  const data = loadData();
  syncDaysToTasks(data);
  syncTopicData(data);
  const todayKey = ensureToday(data);

  todayLabel.textContent = formatDate();
  resetButton.addEventListener("click", () => resetDay(data, todayKey));
  addTaskButton.addEventListener("click", () => addTask(data, todayKey));
  addTopicButton.addEventListener("click", () => addTopic(data, todayKey));
  addMockTestButton.addEventListener("click", () => addMockTest(data, todayKey));
  editorList.addEventListener("input", (event) => handleTaskEditorInput(event, data, todayKey));
  editorList.addEventListener("change", () => refreshAfterEdit(data, todayKey));
  editorList.addEventListener("click", (event) => handleTaskEditorClick(event, data, todayKey));
  weeklyReviewBody.addEventListener("input", (event) => handleTableInput(event, data, todayKey));
  weeklyReviewBody.addEventListener("change", () => refreshAfterEdit(data, todayKey));
  mockTestsBody.addEventListener("input", (event) => handleTableInput(event, data, todayKey));
  mockTestsBody.addEventListener("change", () => refreshAfterEdit(data, todayKey));
  mockTestsBody.addEventListener("click", (event) => handleMockTestClick(event, data, todayKey));
  journalEntry.addEventListener("input", () => handleJournalInput(data, todayKey));
  topicGroups.addEventListener("change", (event) => handleTopicChange(event, data, todayKey));
  topicEditorList.addEventListener("input", (event) => handleTopicEditorInput(event, data, todayKey));
  topicEditorList.addEventListener("change", () => refreshAfterEdit(data, todayKey));
  topicEditorList.addEventListener("click", (event) => handleTopicEditorClick(event, data, todayKey));

  saveData(data);
  updateUI(data, todayKey);
}

init();
