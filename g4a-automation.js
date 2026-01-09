/**
* @OnlyCurrentDoc
* 
* Google Analytics 4 Automated Reporting Script
* Generates comprehensive GA4 reports with period-over-period comparison
* 
* Features:
* - Period-over-period comparison
* - Visual charts (pie, bar)
* - AI-powered insights (optional)
* - Email delivery
* - Automated scheduling
* - Interactive chatbot interface
*/

// ==================== CONFIGURATION ====================
// UPDATE THESE VALUES WITH YOUR OWN DETAILS

const CONFIG = {
  GA4_PROPERTY_ID: 'YOUR_GA4_PROPERTY_ID',           // Your GA4 Property ID (numbers only)
  EMAIL_RECIPIENTS: ['your-email@example.com'],       // Email addresses to receive reports
  GEMINI_API_KEY: 'YOUR_GEMINI_API_KEY',             // Optional: For AI insights
  COMPANY_NAME: 'Your Company Name',                  // Your company/website name
  WEBSITE_URL: 'https://www.yourwebsite.com',        // Your website URL
  GEMINI_MODEL: 'gemini-2.0-flash-exp',              // Gemini model for AI insights
  REPORT_DAYS: 30,                                    // Default reporting period (days)
  INCLUDE_AI_INSIGHTS: true,                          // Enable/disable AI insights
  CONTACT_PAGE_PATH: '/contact-us',                   // Your contact page path
  EXCLUDE_COUNTRIES: ['India'],                       // Countries to exclude from reports
  EXCLUDE_CITIES: ['Ashburn']                         // Cities to exclude from reports
};

// ==================== OAUTH SETUP ====================
/**
 * Run this function first to get OAuth setup instructions
 */
function setupOAuthScopes() {
  console.log('OAUTH SETUP INSTRUCTIONS FOR GA4 REPORTING');
  console.log('==========================================');
  console.log('');
  console.log('STEP 1: ENABLE MANIFEST FILE');
  console.log('1. Go to Project Settings (gear icon)');
  console.log('2. Check "Show appsscript.json manifest file"');
  console.log('3. Click Save');
  console.log('');
  console.log('STEP 2: UPDATE MANIFEST FILE');
  console.log('1. Click on "appsscript.json" in file list');
  console.log('2. Replace content with the following:');
  console.log('');
  console.log('--- COPY BELOW ---');
  console.log('{');
  console.log('  "timeZone": "America/New_York",');
  console.log('  "dependencies": {},');
  console.log('  "exceptionLogging": "STACKDRIVER",');
  console.log('  "runtimeVersion": "V8",');
  console.log('  "oauthScopes": [');
  console.log('    "https://www.googleapis.com/auth/analytics.readonly",');
  console.log('    "https://www.googleapis.com/auth/script.send_mail",');
  console.log('    "https://www.googleapis.com/auth/spreadsheets",');
  console.log('    "https://www.googleapis.com/auth/script.external_request"');
  console.log('  ]');
  console.log('}');
  console.log('--- END COPY ---');
  console.log('');
  console.log('STEP 3: Save and authorize the script');
}

// ==================== MAIN FUNCTIONS ====================

/**
 * Main function to generate and send GA4 report
 * @param {number} days - Number of days to include in report (default: CONFIG.REPORT_DAYS)
 */
function runFocusedReport(days = CONFIG.REPORT_DAYS) {
  try {
    console.log(`Starting ${days}-day GA4 report with period comparison...`);
    
    // Get current and previous period data
    const reportData = getFocusedGA4Data(days);
    
    // Generate AI insights (optional)
    let aiInsights = '';
    if (CONFIG.INCLUDE_AI_INSIGHTS && CONFIG.GEMINI_API_KEY) {
      aiInsights = generateInsights(reportData);
    }
    
    // Send email report
    sendFocusedEmailReport(reportData, aiInsights, days);
    
    console.log(`${days}-day report sent successfully!`);
    
  } catch (error) {
    console.error('Error in runFocusedReport:', error);
    sendErrorNotification(error);
  }
}

// ==================== DATA FETCHING ====================

/**
 * Fetch GA4 data for current and previous periods
 */
function getFocusedGA4Data(days = 30) {
  // Calculate date ranges
  const currentEndDate = new Date();
  currentEndDate.setDate(currentEndDate.getDate() - 1);
  const currentStartDate = new Date(currentEndDate);
  currentStartDate.setDate(currentStartDate.getDate() - days + 1);
  
  const previousEndDate = new Date(currentStartDate);
  previousEndDate.setDate(previousEndDate.getDate() - 1);
  const previousStartDate = new Date(previousEndDate);
  previousStartDate.setDate(previousStartDate.getDate() - days + 1);
  
  const currentPeriod = {
    start: Utilities.formatDate(currentStartDate, 'UTC', 'yyyy-MM-dd'),
    end: Utilities.formatDate(currentEndDate, 'UTC', 'yyyy-MM-dd')
  };
  
  const previousPeriod = {
    start: Utilities.formatDate(previousStartDate, 'UTC', 'yyyy-MM-dd'),
    end: Utilities.formatDate(previousEndDate, 'UTC', 'yyyy-MM-dd')
  };
  
  console.log(`Current period: ${currentPeriod.start} to ${currentPeriod.end}`);
  console.log(`Previous period: ${previousPeriod.start} to ${previousPeriod.end}`);
  
  // Build geographic filter
  const geoFilter = buildGeographicFilter();

  return {
    allChannels: {
      current: getAllChannelsData(currentPeriod.start, currentPeriod.end, geoFilter),
      previous: getAllChannelsData(previousPeriod.start, previousPeriod.end, geoFilter)
    },
    organicSearchPages: getOrganicSearchPages(currentPeriod.start, currentPeriod.end, geoFilter),
    organicSocialData: getOrganicSocialData(currentPeriod.start, currentPeriod.end, geoFilter),
    unassignedSources: getUnassignedSources(currentPeriod.start, currentPeriod.end, geoFilter),
    topRegions: getTopRegions(currentPeriod.start, currentPeriod.end, geoFilter),
    genderData: getGenderData(currentPeriod.start, currentPeriod.end, geoFilter),
    ageData: getAgeData(currentPeriod.start, currentPeriod.end, geoFilter),
    deviceData: getDeviceData(currentPeriod.start, currentPeriod.end, geoFilter),
    contactPageExits: getContactPageExits(currentPeriod.start, currentPeriod.end, geoFilter),
    currentPeriod,
    previousPeriod,
    days,
    filters: buildFilterDescription()
  };
}

/**
 * Build geographic filter based on configuration
 */
function buildGeographicFilter() {
  const expressions = [];
  
  // Exclude countries
  CONFIG.EXCLUDE_COUNTRIES.forEach(country => {
    expressions.push({
      "notExpression": {
        "filter": {
          "fieldName": "country",
          "stringFilter": {
            "matchType": "EXACT",
            "value": country
          }
        }
      }
    });
  });
  
  // Exclude cities
  CONFIG.EXCLUDE_CITIES.forEach(city => {
    expressions.push({
      "notExpression": {
        "filter": {
          "fieldName": "city",
          "stringFilter": {
            "matchType": "CONTAINS",
            "value": city
          }
        }
      }
    });
  });
  
  return { "andGroup": { "expressions": expressions } };
}

/**
 * Build human-readable filter description
 */
function buildFilterDescription() {
  const parts = [];
  if (CONFIG.EXCLUDE_COUNTRIES.length > 0) {
    parts.push(`Excluding countries: ${CONFIG.EXCLUDE_COUNTRIES.join(', ')}`);
  }
  if (CONFIG.EXCLUDE_CITIES.length > 0) {
    parts.push(`Excluding cities: ${CONFIG.EXCLUDE_CITIES.join(', ')}`);
  }
  return parts.join(' and ');
}

// ==================== GA4 API FUNCTIONS ====================

/**
 * Get all channels data
 */
function getAllChannelsData(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "sessionDefaultChannelGrouping" }],
    "metrics": [
      { "name": "sessions" },
      { "name": "newUsers" },
      { "name": "totalUsers" },
      { "name": "averageSessionDuration" },
      { "name": "userEngagementDuration" },
      { "name": "eventsPerSession" },
      { "name": "bounceRate" }
    ],
    "dimensionFilter": geoFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }]
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get organic search landing pages
 */
function getOrganicSearchPages(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const organicSearchFilter = {
    "andGroup": {
      "expressions": [
        ...geoFilter.andGroup.expressions,
        {
          "filter": {
            "fieldName": "sessionDefaultChannelGrouping",
            "stringFilter": { "matchType": "EXACT", "value": "Organic Search" }
          }
        }
      ]
    }
  };
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "landingPage" }],
    "metrics": [
      { "name": "sessions" },
      { "name": "newUsers" },
      { "name": "totalUsers" },
      { "name": "averageSessionDuration" },
      { "name": "userEngagementDuration" },
      { "name": "eventsPerSession" },
      { "name": "bounceRate" }
    ],
    "dimensionFilter": organicSearchFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }],
    "limit": 15
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get organic social data
 */
function getOrganicSocialData(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const organicSocialFilter = {
    "andGroup": {
      "expressions": [
        ...geoFilter.andGroup.expressions,
        {
          "filter": {
            "fieldName": "sessionDefaultChannelGrouping",
            "stringFilter": { "matchType": "EXACT", "value": "Organic Social" }
          }
        }
      ]
    }
  };
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [
      { "name": "landingPage" },
      { "name": "sessionCampaignName" },
      { "name": "sessionSource" }
    ],
    "metrics": [
      { "name": "sessions" },
      { "name": "newUsers" },
      { "name": "totalUsers" },
      { "name": "averageSessionDuration" },
      { "name": "userEngagementDuration" },
      { "name": "eventsPerSession" },
      { "name": "bounceRate" }
    ],
    "dimensionFilter": organicSocialFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }],
    "limit": 5
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get unassigned sources
 */
function getUnassignedSources(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const unassignedFilter = {
    "andGroup": {
      "expressions": [
        ...geoFilter.andGroup.expressions,
        {
          "filter": {
            "fieldName": "sessionDefaultChannelGrouping",
            "stringFilter": { "matchType": "EXACT", "value": "Unassigned" }
          }
        }
      ]
    }
  };
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "firstUserSource" }],
    "metrics": [
      { "name": "sessions" },
      { "name": "newUsers" },
      { "name": "totalUsers" },
      { "name": "averageSessionDuration" },
      { "name": "userEngagementDuration" },
      { "name": "eventsPerSession" },
      { "name": "bounceRate" }
    ],
    "dimensionFilter": unassignedFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }],
    "limit": 10
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get top regions
 */
function getTopRegions(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "country" }],
    "metrics": [{ "name": "sessions" }],
    "dimensionFilter": geoFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }],
    "limit": 5
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get gender data
 */
function getGenderData(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "userGender" }],
    "metrics": [{ "name": "sessions" }],
    "dimensionFilter": geoFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }]
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    console.log('Gender data not available');
    return [];
  }
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get age data
 */
function getAgeData(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "userAgeBracket" }],
    "metrics": [{ "name": "sessions" }],
    "dimensionFilter": geoFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }]
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    console.log('Age data not available');
    return [];
  }
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get device data
 */
function getDeviceData(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "deviceCategory" }],
    "metrics": [
      { "name": "sessions" },
      { "name": "newUsers" },
      { "name": "totalUsers" },
      { "name": "averageSessionDuration" },
      { "name": "userEngagementDuration" },
      { "name": "eventsPerSession" },
      { "name": "bounceRate" }
    ],
    "dimensionFilter": geoFilter,
    "orderBys": [{ "metric": { "metricName": "sessions" }, "desc": true }]
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

/**
 * Get contact page exits
 */
function getContactPageExits(startDate, endDate, geoFilter) {
  const accessToken = ScriptApp.getOAuthToken();
  const apiUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
  
  const contactPageFilter = {
    "andGroup": {
      "expressions": [
        ...geoFilter.andGroup.expressions,
        {
          "filter": {
            "fieldName": "pagePath",
            "stringFilter": {
              "matchType": "EXACT",
              "value": CONFIG.CONTACT_PAGE_PATH
            }
          }
        }
      ]
    }
  };
  
  const payload = {
    "dateRanges": [{ "startDate": startDate, "endDate": endDate }],
    "dimensions": [{ "name": "pagePath" }],
    "metrics": [
      { "name": "sessions" },
      { "name": "screenPageViews" },
      { "name": "totalUsers" }
    ],
    "dimensionFilter": contactPageFilter
  };
  
  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload)
  });
  
  return parseGA4Response(JSON.parse(response.getContentText()));
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * Parse GA4 API response
 */
function parseGA4Response(response) {
  if (!response.rows || response.rows.length === 0) return [];
  
  return response.rows.map(row => {
    const result = {};
    
    if (row.dimensionValues && response.dimensionHeaders) {
      response.dimensionHeaders.forEach((header, index) => {
        result[header.name] = row.dimensionValues[index].value;
      });
    }
    
    if (row.metricValues && response.metricHeaders) {
      response.metricHeaders.forEach((header, index) => {
        let value = row.metricValues[index].value;
        
        if (header.name.includes('Rate') || header.name.includes('Duration') || header.name.includes('Per')) {
          value = parseFloat(value) || 0;
        } else {
          value = parseInt(value) || 0;
        }
        
        result[header.name] = value;
      });
    }
    
    return result;
  });
}

/**
 * Enrich metrics with calculated fields
 */
function enrichMetrics(data) {
  return data.map(row => ({
    ...row,
    returningUsers: (row.totalUsers || 0) - (row.newUsers || 0),
    avgEngagementPerSession: row.sessions > 0 ? (row.userEngagementDuration || 0) / row.sessions : 0
  }));
}

/**
 * Calculate period-over-period changes
 */
function calculatePeriodChanges(current, previous) {
  return current.map(currentRow => {
    const channelName = currentRow.sessionDefaultChannelGrouping;
    const previousRow = previous.find(p => p.sessionDefaultChannelGrouping === channelName) || {};
    
    const calculateChange = (curr, prev) => {
      if (prev === 0) return curr > 0 ? 100 : 0;
      return ((curr - prev) / prev * 100);
    };
    
    return {
      ...enrichMetrics([currentRow])[0],
      changes: {
        sessions: calculateChange(currentRow.sessions || 0, previousRow.sessions || 0),
        newUsers: calculateChange(currentRow.newUsers || 0, previousRow.newUsers || 0),
        totalUsers: calculateChange(currentRow.totalUsers || 0, previousRow.totalUsers || 0),
        averageSessionDuration: calculateChange(currentRow.averageSessionDuration || 0, previousRow.averageSessionDuration || 0),
        avgEngagementPerSession: calculateChange(
          (currentRow.userEngagementDuration || 0) / (currentRow.sessions || 1),
          (previousRow.userEngagementDuration || 0) / (previousRow.sessions || 1)
        ),
        eventsPerSession: calculateChange(currentRow.eventsPerSession || 0, previousRow.eventsPerSession || 0),
        bounceRate: calculateChange(currentRow.bounceRate || 0, previousRow.bounceRate || 0)
      }
    };
  });
}

// ==================== CHART GENERATION ====================

/**
 * Generate CSS-based pie chart
 */
function generatePieChart(data, title, totalSessions = null) {
  if (!data || data.length === 0) {
    return `<div style="text-align: center; padding: 40px; color: #666;">No ${title} data available</div>`;
  }
  
  const total = totalSessions || data.reduce((sum, item) => sum + item.sessions, 0);
  const colors = ['#1a73e8', '#34a853', '#fbbc05', '#ea4335', '#9aa0a6'];
  
  let cumulativePercent = 0;
  const segments = data.map((item, index) => {
    const percent = (item.sessions / total) * 100;
    const startAngle = cumulativePercent * 3.6;
    cumulativePercent += percent;
    const endAngle = cumulativePercent * 3.6;
    
    const largeArcFlag = percent > 50 ? 1 : 0;
    const x1 = 50 + 40 * Math.cos((startAngle - 90) * Math.PI / 180);
    const y1 = 50 + 40 * Math.sin((startAngle - 90) * Math.PI / 180);
    const x2 = 50 + 40 * Math.cos((endAngle - 90) * Math.PI / 180);
    const y2 = 50 + 40 * Math.sin((endAngle - 90) * Math.PI / 180);
    
    const pathData = [
      `M 50 50`,
      `L ${x1} ${y1}`,
      `A 40 40 0 ${largeArcFlag} 1 ${x2} ${y2}`,
      'Z'
    ].join(' ');
    
    return {
      path: pathData,
      color: colors[index % colors.length],
      label: item[Object.keys(item)[0]] || 'Unknown',
      value: item.sessions,
      percent: percent.toFixed(1)
    };
  });
  
  const svgSegments = segments.map(segment => 
    `<path d="${segment.path}" fill="${segment.color}" stroke="white" stroke-width="1"/>`
  ).join('');
  
  const legend = segments.map(segment => 
    `<div style="display: flex; align-items: center; margin: 5px 0;">
      <div style="width: 12px; height: 12px; background: ${segment.color}; margin-right: 8px; border-radius: 2px;"></div>
      <span style="font-size: 12px;">${segment.label}: ${segment.value} (${segment.percent}%)</span>
    </div>`
  ).join('');
  
  return `
    <div style="display: flex; align-items: center; justify-content: center; gap: 30px;">
      <svg width="120" height="120" viewBox="0 0 100 100">
        ${svgSegments}
      </svg>
      <div style="font-size: 14px;">
        ${legend}
      </div>
    </div>
  `;
}

/**
 * Generate CSS-based bar chart
 */
function generateBarChart(data, title) {
  if (!data || data.length === 0) {
    return `<div style="text-align: center; padding: 40px; color: #666;">No ${title} data available</div>`;
  }
  
  const maxValue = Math.max(...data.map(item => item.sessions));
  
  const bars = data.map((item, index) => {
    const percentage = (item.sessions / maxValue) * 100;
    const label = item[Object.keys(item)[0]] || 'Unknown';
    
    return `
      <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <div style="width: 80px; font-size: 12px; text-align: right; padding-right: 10px;">${label}:</div>
        <div style="flex: 1; background: #f0f0f0; height: 24px; border-radius: 4px; position: relative;">
          <div style="background: #1a73e8; height: 100%; width: ${percentage}%; border-radius: 4px; display: flex; align-items: center; justify-content: flex-end; padding-right: 5px;">
            <span style="color: white; font-size: 11px; font-weight: bold;">${item.sessions}</span>
          </div>
        </div>
      </div>
    `;
  }).join('');
  
  return `<div style="max-width: 400px;">${bars}</div>`;
}

// ==================== AI INSIGHTS ====================

/**
 * Generate AI insights using Gemini
 */
function generateInsights(reportData) {
  if (!CONFIG.GEMINI_API_KEY) {
    return getFallbackInsights(reportData);
  }

  const currentChannels = reportData.allChannels.current;
  const totalSessions = currentChannels.reduce((sum, channel) => sum + channel.sessions, 0);
  
  const top5Pages = reportData.organicSearchPages.slice(0, 5).map(page => {
    return `${page.landingPage}: ${page.sessions.toLocaleString()} sessions`;
  }).join('\n');
  
  const channelSummary = currentChannels.map(channel => {
    const percentage = ((channel.sessions / totalSessions) * 100).toFixed(1);
    return `${channel.sessionDefaultChannelGrouping}: ${channel.sessions.toLocaleString()} sessions (${percentage}%)`;
  }).join('\n');
  
  const prompt = `Provide a positive summary for this ${reportData.days}-day analytics data:

TOTAL TRAFFIC: ${totalSessions.toLocaleString()} sessions

TOP 5 PAGES:
${top5Pages}

CHANNEL-WISE SUMMARY:
${channelSummary}

Write a concise positive overview highlighting the performance metrics. Do not include suggestions or recommendations. Focus only on summarizing the data in a positive manner.`;

  try {
    const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;
    
    const response = UrlFetchApp.fetch(geminiUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0.3,
          maxOutputTokens: 200
        }
      })
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts) {
      let insights = data.candidates[0].content.parts[0].text.trim();
      insights = insights.replace(/\/n/g, '<br>');
      return insights;
    }
    
  } catch (error) {
    console.error('Error generating insights:', error);
  }
  
  return getFallbackInsights(reportData);
}

/**
 * Fallback insights when AI is not available
 */
function getFallbackInsights(reportData) {
  const currentChannels = reportData.allChannels.current;
  const totalSessions = currentChannels.reduce((sum, channel) => sum + channel.sessions, 0);
  
  return `${CONFIG.COMPANY_NAME} achieved ${totalSessions.toLocaleString()} total sessions over the ${reportData.days}-day period.<br>Top performing channel: ${currentChannels[0]?.sessionDefaultChannelGrouping || 'N/A'} with strong engagement metrics.<br>Organic search continues to drive quality traffic with ${reportData.organicSearchPages.length} active landing pages.`;
}

// ==================== EMAIL REPORT ====================

/**
 * Send focused email report
 */
function sendFocusedEmailReport(reportData, aiInsights, days) {
  const subject = `${days}-Day Analytics Report - ${CONFIG.COMPANY_NAME}`;
  
  const channelsWithChanges = calculatePeriodChanges(reportData.allChannels.current, reportData.allChannels.previous);
  const organicSearchEnriched = enrichMetrics(reportData.organicSearchPages);
  const organicSocialEnriched = enrichMetrics(reportData.organicSocialData);
  const unassignedEnriched = enrichMetrics(reportData.unassignedSources);
  const deviceEnriched = enrichMetrics(reportData.deviceData);
  
  const regionsPieChart = generatePieChart(reportData.topRegions, 'regions');
  const genderPieChart = generatePieChart(reportData.genderData, 'gender');
  const ageBarChart = generateBarChart(reportData.ageData, 'age groups');
  
  const formatDuration = (seconds) => {
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = Math.floor(seconds % 60);
    return `${minutes}m ${remainingSeconds}s`;
  };
  
  const formatChange = (change) => {
    const arrow = change > 0 ? '↗️' : change < 0 ? '↘️' : '→';
    const color = change > 0 ? '#137333' : change < 0 ? '#ea4335' : '#5f6368';
    return `<span style="color: ${color}; font-weight: 600;">${arrow} ${Math.abs(change).toFixed(1)}%</span>`;
  };
  
  const htmlBody = `
    <html>
      <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; line-height: 1.6; color: #202124; background: #f8f9fa; margin: 0; padding: 20px;">
        <div style="max-width: 1200px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden;">
          
          <!-- Header -->
          <div style="background: linear-gradient(135deg, #1a73e8 0%, #4285f4 100%); color: white; padding: 30px; text-align: center;">
            <h1 style="margin: 0; font-size: 28px; font-weight: 500;">${days}-Day Analytics Report</h1>
            <p style="margin: 10px 0 0 0; opacity: 0.9; font-size: 16px;">${CONFIG.COMPANY_NAME}</p>
            <p style="margin: 5px 0 0 0; opacity: 0.8; font-size: 14px;">Current: ${reportData.currentPeriod.start} to ${reportData.currentPeriod.end}</p>
            <p style="margin: 0; opacity: 0.8; font-size: 14px;">vs Previous: ${reportData.previousPeriod.start} to ${reportData.previousPeriod.end}</p>
            ${reportData.filters ? `<p style="margin: 5px 0 0 0; opacity: 0.8; font-size: 12px;">${reportData.filters}</p>` : ''}
          </div>
          
          <div style="padding: 30px;">
            
            ${CONFIG.INCLUDE_AI_INSIGHTS && aiInsights ? `
            <!-- AI Insights -->
            <div style="background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); padding: 20px; border-radius: 12px; margin-bottom: 30px; border-left: 4px solid #9c27b0;">
              <h2 style="margin: 0 0 15px 0; color: #6a1b9a; font-size: 18px;">AI-Powered Insights</h2>
              <p style="margin: 0; font-size: 14px; line-height: 1.6; color: #4a148c;">${aiInsights}</p>
            </div>
            ` : ''}
            
            <!-- 1. All Channels Comparison -->
            <h2 style="color: #1565c0; font-size: 20px; margin: 0 0 20px 0; border-bottom: 2px solid #e3f2fd; padding-bottom: 8px;">1. Website Traffic by Channel</h2>
            
            <div style="overflow-x: auto; margin-bottom: 30px;">
              <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                <thead>
                  <tr style="background: #f8f9fa; border-bottom: 2px solid #e8eaed;">
                    <th style="padding: 12px; text-align: left; font-weight: 600; color: #3c4043;">Channel</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Sessions</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">New Users</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Returning Users</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Avg Duration</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Engagement/Session</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Events/Session</th>
                    <th style="padding: 12px; text-align: right; font-weight: 600; color: #3c4043;">Bounce Rate</th>
                  </tr>
                </thead>
                <tbody>
                  ${channelsWithChanges.map((channel, index) => {
                    const rowColor = index % 2 === 0 ? '#fafbfc' : 'white';
                    return `
                    <tr style="background: ${rowColor}; border-bottom: 1px solid #e8eaed;">
                      <td style="padding: 12px; font-weight: 600; color: #1a73e8;">${channel.sessionDefaultChannelGrouping}</td>
                      <td style="padding: 12px; text-align: right;">
                        <div style="font-weight: 600;">${channel.sessions.toLocaleString()}</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.sessions)}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div>${channel.newUsers.toLocaleString()}</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.newUsers)}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div>${channel.returningUsers.toLocaleString()}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div>${formatDuration(channel.averageSessionDuration)}</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.averageSessionDuration)}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div>${channel.avgEngagementPerSession.toFixed(0)}s</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.avgEngagementPerSession)}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div>${channel.eventsPerSession.toFixed(2)}</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.eventsPerSession)}</div>
                      </td>
                      <td style="padding: 12px; text-align: right;">
                        <div style="color: ${(channel.bounceRate * 100) < 60 ? '#137333' : '#ea4335'};">${(channel.bounceRate * 100).toFixed(1)}%</div>
                        <div style="font-size: 11px;">${formatChange(channel.changes.bounceRate)}</div>
                      </td>
                    </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
            
            <!-- Visual Charts Section -->
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 30px; margin-bottom: 30px;">
              <div>
                <h2 style="color: #1565c0; font-size: 20px; margin: 0 0 20px 0; border-bottom: 2px solid #e3f2fd; padding-bottom: 8px;">Top 5 Regions</h2>
                <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center;">
                  ${regionsPieChart}
                </div>
              </div>
              
              <div>
                <h2 style="color: #1565c0; font-size: 20px; margin: 0 0 20px 0; border-bottom: 2px solid #e3f2fd; padding-bottom: 8px;">Gender Distribution</h2>
                <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center;">
                  ${genderPieChart}
                </div>
              </div>
            </div>
            
            <!-- Summary -->
            <div style="background: #e8f0fe; border-radius: 12px; padding: 25px; margin: 30px 0 0 0;">
              <h3 style="margin: 0 0 15px 0; color: #1565c0; font-size: 18px;">${days}-Day Summary</h3>
              <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px;">
                <div style="text-align: center;">
                  <div style="font-size: 14px; color: #5f6368;">Total Sessions</div>
                  <div style="font-size: 28px; font-weight: 700; color: #1565c0;">${channelsWithChanges.reduce((sum, ch) => sum + ch.sessions, 0).toLocaleString()}</div>
                </div>
                <div style="text-align: center;">
                  <div style="font-size: 14px; color: #5f6368;">Top Channel</div>
                  <div style="font-size: 16px; font-weight: 600; color: #1565c0;">${channelsWithChanges[0]?.sessionDefaultChannelGrouping || 'N/A'}</div>
                </div>
                <div style="text-align: center;">
                  <div style="font-size: 14px; color: #5f6368;">Period Comparison</div>
                  <div style="font-size: 16px; font-weight: 600; color: #1565c0;">vs Previous ${days} Days</div>
                </div>
              </div>
            </div>
            
          </div>
        </div>
      </body>
    </html>
  `;
  
  CONFIG.EMAIL_RECIPIENTS.forEach(email => {
    GmailApp.sendEmail(
      email,
      subject,
      '',
      {
        htmlBody: htmlBody,
        name: CONFIG.COMPANY_NAME + ' Analytics Report'
      }
    );
  });
  
  console.log(`Report sent to: ${CONFIG.EMAIL_RECIPIENTS.join(', ')}`);
}

// ==================== ERROR HANDLING ====================

/**
 * Send error notification
 */
function sendErrorNotification(error) {
  const subject = `GA4 Report Error - ${CONFIG.COMPANY_NAME}`;
  const body = `
An error occurred while generating the GA4 report:

Error: ${error.toString()}
Time: ${new Date().toISOString()}

Please check the script configuration and try again.

Common issues:
1. GA4 API permissions - check OAuth scopes
2. Invalid Property ID: ${CONFIG.GA4_PROPERTY_ID}
3. Gemini API key issues (if AI insights enabled)
4. Demographics data not available

Please review your configuration and run enableGA4API() to test connectivity.
  `;
  
  CONFIG.EMAIL_RECIPIENTS.forEach(email => {
    GmailApp.sendEmail(email, subject, body);
  });
}

// ==================== TESTING & SETUP ====================

/**
 * Test GA4 API connection
 */
function enableGA4API() {
  console.log('Testing GA4 API authentication...');
  
  try {
    const accessToken = ScriptApp.getOAuthToken();
    console.log('OAuth token obtained');
    
    const testUrl = `https://analyticsdata.googleapis.com/v1beta/properties/${CONFIG.GA4_PROPERTY_ID}:runReport`;
    
    const testPayload = {
      "dateRanges": [{"startDate": "7daysAgo", "endDate": "yesterday"}],
      "metrics": [{"name": "sessions"}],
      "limit": 1
    };
    
    const response = UrlFetchApp.fetch(testUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(testPayload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const sessions = data.rows?.[0]?.metricValues?.[0]?.value || 'No data';
      console.log('✓ GA4 API access successful!');
      console.log(`Test result: ${sessions} sessions in last 7 days`);
      console.log('Ready to generate reports!');
    } else {
      console.log('✗ API Error Code:', responseCode);
      console.log('Response:', response.getContentText());
      
      if (responseCode === 403) {
        console.log('Solution: Run setupOAuthScopes() and configure OAuth permissions');
      } else if (responseCode === 400) {
        console.log(`Solution: Verify GA4 Property ID: ${CONFIG.GA4_PROPERTY_ID}`);
      }
    }
    
  } catch (error) {
    console.error('GA4 API test failed:', error.toString());
  }
}

/**
 * Test report generation
 */
function testFocusedReport() {
  console.log('Testing report generation with 7-day period...');
  try {
    runFocusedReport(7);
    console.log('✓ Test report completed successfully!');
  } catch (error) {
    console.error('✗ Test report failed:', error);
  }
}

/**
 * System check
 */
function runSystemCheck() {
  console.log('GA4 AUTOMATION SYSTEM CHECK');
  console.log('===========================');
  console.log('');
  
  console.log('Configuration:');
  console.log(`- Property ID: ${CONFIG.GA4_PROPERTY_ID}`);
  console.log(`- Recipients: ${CONFIG.EMAIL_RECIPIENTS.join(', ')}`);
  console.log(`- Default Period: ${CONFIG.REPORT_DAYS} days`);
  console.log(`- AI Insights: ${CONFIG.INCLUDE_AI_INSIGHTS ? 'Enabled' : 'Disabled'}`);
  console.log(`- Company: ${CONFIG.COMPANY_NAME}`);
  console.log(`- Website: ${CONFIG.WEBSITE_URL}`);
  console.log('');
  
  console.log('Testing GA4 API access...');
  enableGA4API();
}

// ==================== AUTOMATION ====================

/**
 * Setup automated scheduling
 */
function setupAutomation(frequencyDays = 7, reportDays = CONFIG.REPORT_DAYS) {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runScheduledFocusedReport') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger
  ScriptApp.newTrigger('runScheduledFocusedReport')
    .timeBased()
    .everyDays(frequencyDays)
    .atHour(9)
    .create();
  
  PropertiesService.getScriptProperties().setProperties({
    'SCHEDULED_REPORT_DAYS': reportDays.toString(),
    'SCHEDULED_FREQUENCY': frequencyDays.toString()
  });
  
  console.log(`✓ Automation setup: ${reportDays}-day report every ${frequencyDays} days at 9 AM`);
}

/**
 * Scheduled report handler
 */
function runScheduledFocusedReport() {
  const reportDays = parseInt(PropertiesService.getScriptProperties().getProperty('SCHEDULED_REPORT_DAYS') || CONFIG.REPORT_DAYS);
  runFocusedReport(reportDays);
}

// ==================== CONVENIENCE FUNCTIONS ====================

function runWeeklyReport() {
  runFocusedReport(7);
}

function runMonthlyReport() {
  runFocusedReport(30);
}

function run90DayReport() {
  runFocusedReport(90);
}
