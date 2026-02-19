// helpers/apiCache.js
import { API_CONFIG, SETTINGS } from './constants.js';

class SheetsAPICache {
  constructor() {
    this.cache = new Map();
    this.pendingRequests = new Map();
  }

  _getCacheKey(sheetId, ranges) {
    return `${sheetId}:${[...ranges].sort().join(',')}`;
  }

  _isExpired(timestamp) {
    return Date.now() - timestamp > API_CONFIG.CACHE_TTL;
  }

  async batchGet(sheetId, ranges) {
    const cacheKey = this._getCacheKey(sheetId, ranges);
    
    // Check if we have a pending request for the same data
    if (this.pendingRequests.has(cacheKey)) {
      return await this.pendingRequests.get(cacheKey);
    }

    // Check cache first
    const cached = this.cache.get(cacheKey);
    if (cached && !this._isExpired(cached.timestamp)) {
      return cached.data;
    }

    // Create the API request promise
    const requestPromise = this._fetchFromAPI(sheetId, ranges);
    this.pendingRequests.set(cacheKey, requestPromise);

    try {
      const data = await requestPromise;
      
      // Cache the result
      this.cache.set(cacheKey, {
        data,
        timestamp: Date.now()
      });

      return data;
    } catch (error) {
      console.error("❌ Sheexcel | API request failed:", error);
      
      // Return stale data if available
      if (cached) {
        ui.notifications.warn("Using cached data - connection issues detected");
        return cached.data;
      }
      
      throw error;
    } finally {
      this.pendingRequests.delete(cacheKey);
    }
  }

  async _fetchFromAPI(sheetId, ranges) {
    const apiKey = game.settings.get("sheexcel_updated", SETTINGS.GOOGLE_API_KEY);
    if (!apiKey) {
      throw new Error("Google API key not configured");
    }

    const rangesParam = ranges.map(r => `ranges=${encodeURIComponent(r)}`).join("&");
    const url = `${API_CONFIG.BASE_URL}/${sheetId}/values:batchGet?key=${apiKey}&${rangesParam}`;

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), API_CONFIG.TIMEOUT);

    try {
      const response = await fetch(url, { 
        signal: controller.signal,
        headers: {
          'Accept': 'application/json'
        }
      });
      
      clearTimeout(timeoutId);

      if (!response.ok) {
        throw new Error(`API Error: ${response.status} ${response.statusText}`);
      }

      return await response.json();
    } catch (error) {
      clearTimeout(timeoutId);
      
      if (error.name === 'AbortError') {
        throw new Error('Request timeout - check your connection');
      }
      
      throw error;
    }
  }

  invalidateSheet(sheetId) {
    // Remove all cached entries for this sheet
    for (const [key] of this.cache) {
      if (key.startsWith(`${sheetId}:`)) {
        this.cache.delete(key);
      }
    }
  }

  clearCache() {
    this.cache.clear();
    this.pendingRequests.clear();
  }
}

// Export singleton instance
export const apiCache = new SheetsAPICache();