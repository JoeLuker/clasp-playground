/**
 * Configuration and Constants
 * Central location for all game constants and configuration values.
 */

const SHEET_NAMES = {
  DASHBOARD: 'Dashboard',
  CALCULATIONS: 'Calculations',
  ROSTER: 'Roster',
  STATUS: 'Status Monitor',
  CONSTANTS: 'Constants',
  REFERENCES: 'Reference Tables',
  CARAVAN: 'Caravan',
  EXPLORATION: 'Exploration',
  LOG: 'Log',
  HEX_MAP: 'Hex Map',
  EVENT_LOG: 'Event Log'
};

const SHEET_COLORS = {
  DASHBOARD: '#4285f4',
  CALCULATIONS: '#0f9d58',
  ROSTER: '#ea4335',
  STATUS: '#fbbc04',
  CONSTANTS: '#9c27b0',
  REFERENCES: '#9e9e9e',
  CARAVAN: '#ff6f00',
  EXPLORATION: '#00897b',
  LOG: '#673ab7',
  HEX_MAP: '#00bcd4',
  EVENT_LOG: '#9c27b0'
};

const ROW_OFFSETS = {
  PARTY_START: 4,
  PARTY_COUNT: 4,
  MOUNT_START: 19,
  FORMULA_START: 5
};

const BATCH_SIZE = 1000;

const VALID_TEMPERATURES = ['Extreme Cold', 'Severe Cold', 'Cold', 'Normal', 'Hot', 'Severe Heat', 'Extreme Heat'];
const VALID_TERRAINS = ['Desert, sandy', 'Forest', 'Hills', 'Jungle', 'Moor', 'Mountains', 'Plains', 'Swamp', 'Tundra, frozen'];
const VALID_PATH_TYPES = ['Highway', 'Road or Trail', 'Trackless'];
const VALID_WEATHER = ['Clear', 'Light Rain', 'Heavy Rain', 'Storm', 'Snow', 'Blizzard'];
const VALID_PACES = ['Slow', 'Normal', 'Fast'];
