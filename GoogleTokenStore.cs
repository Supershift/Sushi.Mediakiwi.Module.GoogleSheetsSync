using Google.Apis.Util.Store;
using Sushi.Mediakiwi.Data;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync
{
    public class GoogleTokenStore : IDataStore
    {

        private IApplicationUser _user { get; set; }
        private string _storeIdentifier { get; set; }

        public GoogleTokenStore(IApplicationUser user, string identifier)
        {
            _user = user;
            _storeIdentifier = identifier;
        }

        public async Task ClearAsync()
        {
            foreach (var item in _user.Data.Items)
            {
                if (item.Property.StartsWith(_storeIdentifier, StringComparison.InvariantCulture))
                {
                    item.Apply(null);
                }
            }

            await _user.SaveAsync();
        }

        public async Task DeleteAsync<T>(string key)
        {
            var genKey = GenerateStoredKey(key, typeof(T));

            if (_user?.Data?.HasProperty(genKey) == true)
            {
                _user.Data.Apply(genKey, null);
                await _user.SaveAsync();
            }
        }

        public async Task<T> GetAsync<T>(string key)
        {
            var genKey = GenerateStoredKey(key, typeof(T));

            if (_user?.Data?.HasProperty(genKey) == true)
            {
                return System.Text.Json.JsonSerializer.Deserialize<T>(_user.Data[genKey].Value);
            }
            return default(T);
        }

        public async Task StoreAsync<T>(string key, T value)
        {
            var genKey = GenerateStoredKey(key, typeof(T));

            _user.Data.Apply(genKey, System.Text.Json.JsonSerializer.Serialize(value));

            await _user.SaveAsync();
        }

        private string GenerateStoredKey(string key, Type t)
        {
            return $"{_storeIdentifier}-{t.FullName}-{key}";
        }
    }
}
