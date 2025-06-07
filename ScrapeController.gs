const CANCEL_FLAG_KEY = 'SCRAPE_CANCELLED';

function startScrape() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Could not start scraping: another scrape is already in progress.');
  }
  const cache = CacheService.getUserCache();
  try {
    cache.put(CANCEL_FLAG_KEY, 'false', 3600);
    const summary = ScraperLogic.runScrape();
    return summary;
  } finally {
    cache.remove(CANCEL_FLAG_KEY);
    lock.releaseLock();
  }
}

function cancelScrape() {
  CacheService.getUserCache().put(CANCEL_FLAG_KEY, 'true', 3600);
}

function isCancelled() {
  return CacheService.getUserCache().get(CANCEL_FLAG_KEY) === 'true';
}