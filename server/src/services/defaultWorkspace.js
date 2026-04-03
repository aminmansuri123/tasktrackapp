function defaultWorkspaceData() {
  return {
    users: [],
    tasks: [],
    locations: [
      { id: 1, name: 'Mundra' },
      { id: 2, name: 'JNPT' },
      { id: 3, name: 'Combine' },
    ],
    segregationTypes: [
      { id: 1, name: 'PSA Reports' },
      { id: 2, name: 'Internal Reports' },
    ],
    holidays: [],
    notes: [],
    learningNotes: [],
    milestones: [],
    dailyPlanner: [],
    locationItems: [],
    codeSnippets: [],
    journal: {},
    reportToOptions: [],
  };
}

function normalizeWorkspacePayload(data) {
  if (!data || typeof data !== 'object') {
    return defaultWorkspaceData();
  }
  return {
    users: Array.isArray(data.users) ? data.users : [],
    tasks: Array.isArray(data.tasks) ? data.tasks : [],
    locations: Array.isArray(data.locations) ? data.locations : [],
    segregationTypes: Array.isArray(data.segregationTypes) ? data.segregationTypes : [],
    holidays: Array.isArray(data.holidays) ? data.holidays : [],
    notes: Array.isArray(data.notes) ? data.notes : [],
    learningNotes: Array.isArray(data.learningNotes) ? data.learningNotes : [],
    milestones: Array.isArray(data.milestones) ? data.milestones : [],
    dailyPlanner: Array.isArray(data.dailyPlanner) ? data.dailyPlanner : [],
    locationItems: Array.isArray(data.locationItems) ? data.locationItems : [],
    codeSnippets: Array.isArray(data.codeSnippets) ? data.codeSnippets : [],
    journal: data.journal && typeof data.journal === 'object' ? data.journal : {},
    reportToOptions: Array.isArray(data.reportToOptions) ? data.reportToOptions : [],
  };
}

module.exports = { defaultWorkspaceData, normalizeWorkspacePayload };
