import { UserSettings, DEFAULT_SETTINGS } from "./types";

const SETTINGS_KEY = "quickcel_settings";

export function loadSettings(): UserSettings {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<UserSettings>;
    return { ...DEFAULT_SETTINGS, ...parsed };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: UserSettings): void {
  localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
}

let _cachedSettings: UserSettings | null = null;

export function getCachedSettings(): UserSettings {
  if (!_cachedSettings) {
    _cachedSettings = loadSettings();
  }
  return _cachedSettings;
}

export function refreshSettingsCache(): void {
  _cachedSettings = loadSettings();
}
