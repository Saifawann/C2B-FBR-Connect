using C2B_FBR_Connect.Services;
using System;
using System.Collections.Generic;

namespace C2B_FBR_Connect.Services
{
    /// <summary>
    /// Session-scoped in-memory cache for QuickBooks data
    /// Cache is automatically cleared between invoice processing sessions
    /// Supports manual refresh when QuickBooks data changes
    /// </summary>
    public class DataCacheService
    {
        // Customer cache: ListID -> CustomerData
        private readonly Dictionary<string, CustomerData> _customerCache = new Dictionary<string, CustomerData>();

        // Item cache: ListID -> ItemData
        private readonly Dictionary<string, ItemData> _itemCache = new Dictionary<string, ItemData>();

        // Price levels cache
        private List<PriceLevel> _priceLevelsCache = null;

        // Cache metadata
        private DateTime _cacheCreatedTime = DateTime.Now;
        private readonly Guid _sessionId = Guid.NewGuid();

        // Statistics
        public int CustomerCacheHits { get; private set; }
        public int CustomerCacheMisses { get; private set; }
        public int ItemCacheHits { get; private set; }
        public int ItemCacheMisses { get; private set; }

        #region Customer Cache

        /// <summary>
        /// Try to get customer from cache
        /// </summary>
        public bool TryGetCustomer(string listID, out CustomerData customer)
        {
            if (string.IsNullOrEmpty(listID))
            {
                customer = null;
                return false;
            }

            if (_customerCache.TryGetValue(listID, out customer))
            {
                CustomerCacheHits++;
                return true;
            }

            CustomerCacheMisses++;
            customer = null;
            return false;
        }

        /// <summary>
        /// Add customer to cache
        /// </summary>
        public void AddCustomer(string listID, CustomerData customer)
        {
            if (string.IsNullOrEmpty(listID) || customer == null)
                return;

            _customerCache[listID] = customer;
        }

        /// <summary>
        /// Add multiple customers to cache (batch)
        /// </summary>
        public void AddCustomers(Dictionary<string, CustomerData> customers)
        {
            if (customers == null) return;

            foreach (var kvp in customers)
            {
                _customerCache[kvp.Key] = kvp.Value;
            }
        }

        #endregion

        #region Item Cache

        /// <summary>
        /// Try to get item from cache
        /// </summary>
        public bool TryGetItem(string listID, out ItemData item)
        {
            if (string.IsNullOrEmpty(listID))
            {
                item = null;
                return false;
            }

            if (_itemCache.TryGetValue(listID, out item))
            {
                ItemCacheHits++;
                return true;
            }

            ItemCacheMisses++;
            item = null;
            return false;
        }

        /// <summary>
        /// Add item to cache
        /// </summary>
        public void AddItem(string listID, ItemData item)
        {
            if (string.IsNullOrEmpty(listID) || item == null)
                return;

            _itemCache[listID] = item;
        }

        /// <summary>
        /// Add multiple items to cache (batch)
        /// </summary>
        public void AddItems(Dictionary<string, ItemData> items)
        {
            if (items == null) return;

            foreach (var kvp in items)
            {
                _itemCache[kvp.Key] = kvp.Value;
            }
        }

        #endregion

        #region Price Levels Cache

        /// <summary>
        /// Try to get price levels from cache
        /// </summary>
        public bool TryGetPriceLevels(out List<PriceLevel> priceLevels)
        {
            if (_priceLevelsCache != null)
            {
                priceLevels = _priceLevelsCache;
                return true;
            }

            priceLevels = null;
            return false;
        }

        /// <summary>
        /// Add price levels to cache
        /// </summary>
        public void SetPriceLevels(List<PriceLevel> priceLevels)
        {
            _priceLevelsCache = priceLevels;
        }

        #endregion

        #region Cache Management

        /// <summary>
        /// Clear all caches - call this when QuickBooks data has been modified
        /// </summary>
        public void ClearAll()
        {
            _customerCache.Clear();
            _itemCache.Clear();
            _priceLevelsCache = null;
            _cacheCreatedTime = DateTime.Now;

            // Reset statistics
            CustomerCacheHits = 0;
            CustomerCacheMisses = 0;
            ItemCacheHits = 0;
            ItemCacheMisses = 0;

            System.Diagnostics.Debug.WriteLine($"🔄 Cache cleared - Session: {_sessionId}");
        }

        /// <summary>
        /// Clear only customer cache
        /// </summary>
        public void ClearCustomers()
        {
            _customerCache.Clear();
            CustomerCacheHits = 0;
            CustomerCacheMisses = 0;
            System.Diagnostics.Debug.WriteLine($"🔄 Customer cache cleared");
        }

        /// <summary>
        /// Clear only item cache
        /// </summary>
        public void ClearItems()
        {
            _itemCache.Clear();
            _priceLevelsCache = null;
            ItemCacheHits = 0;
            ItemCacheMisses = 0;
            System.Diagnostics.Debug.WriteLine($"🔄 Item cache cleared");
        }

        /// <summary>
        /// Get cache statistics for monitoring
        /// </summary>
        public string GetCacheStats()
        {
            int totalCustomerRequests = CustomerCacheHits + CustomerCacheMisses;
            int totalItemRequests = ItemCacheHits + ItemCacheMisses;

            double customerHitRate = totalCustomerRequests > 0
                ? (double)CustomerCacheHits / totalCustomerRequests * 100
                : 0;

            double itemHitRate = totalItemRequests > 0
                ? (double)ItemCacheHits / totalItemRequests * 100
                : 0;

            return $@"
📊 Cache Statistics (Session: {_sessionId})
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Created: {_cacheCreatedTime:yyyy-MM-dd HH:mm:ss}
Age: {(DateTime.Now - _cacheCreatedTime).TotalMinutes:F1} minutes

Customers:
  - Cached: {_customerCache.Count}
  - Hits: {CustomerCacheHits}
  - Misses: {CustomerCacheMisses}
  - Hit Rate: {customerHitRate:F1}%

Items:
  - Cached: {_itemCache.Count}
  - Hits: {ItemCacheHits}
  - Misses: {ItemCacheMisses}
  - Hit Rate: {itemHitRate:F1}%

Price Levels: {(_priceLevelsCache != null ? "Cached" : "Not cached")}
";
        }

        /// <summary>
        /// Check if cache is empty (useful for determining if batch load is needed)
        /// </summary>
        public bool IsEmpty()
        {
            return _customerCache.Count == 0 && _itemCache.Count == 0;
        }

        /// <summary>
        /// Get detailed cache contents for debugging
        /// </summary>
        public string GetCacheContents()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"📋 Cache Contents (Session: {_sessionId})");
            sb.AppendLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

            sb.AppendLine($"\n👥 Customers ({_customerCache.Count}):");
            foreach (var kvp in _customerCache)
            {
                sb.AppendLine($"  - {kvp.Key}: {kvp.Value.CustomerType} | NTN: {kvp.Value.NTN ?? "N/A"}");
            }

            sb.AppendLine($"\n📦 Items ({_itemCache.Count}):");
            foreach (var kvp in _itemCache)
            {
                sb.AppendLine($"  - {kvp.Key}: {kvp.Value.SaleType} | HS: {kvp.Value.HSCode ?? "N/A"}");
            }

            sb.AppendLine($"\n💰 Price Levels: {(_priceLevelsCache != null ? $"Cached ({_priceLevelsCache.Count} levels)" : "Not cached")}");

            return sb.ToString();
        }

        #endregion
    }
}