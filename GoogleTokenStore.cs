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
            if (_user?.Data?.HasProperty(_storeIdentifier)==true)
            {
                _user.Data.Apply(_storeIdentifier, null);
                await _user.SaveAsync();
            }
        }

        public async Task DeleteAsync<T>(string key)
        {
            if (_user?.Data?.HasProperty(_storeIdentifier) == true)
            {
                Dictionary<string, T> values = new Dictionary<string, T>();
                var sets = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, T>>(_user.Data[_storeIdentifier].Value);
                if (sets != null)
                {
                    values = sets;
                }

                if (values?.ContainsKey(key) == true)
                { 
                    values.Remove(key);
                    _user.Data.Apply(_storeIdentifier, System.Text.Json.JsonSerializer.Serialize(values));
                    await _user.SaveAsync();
                }
            }
        }

        public async Task<T> GetAsync<T>(string key)
        {
            if (_user?.Data?.HasProperty(_storeIdentifier) == true)
            {
                Dictionary<string, T> values = new Dictionary<string, T>();
                var sets = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, T>>(_user.Data[_storeIdentifier].Value);
                if (sets != null)
                {
                    values = sets;
                }

                if (values?.ContainsKey(key) == true)
                {
                    return values[key];
                }
            }
            return default(T);
        }

        public async Task StoreAsync<T>(string key, T value)
        {
            Dictionary<string, T> values = new Dictionary<string, T>();

            if (_user?.Data?.HasProperty(_storeIdentifier) == true)
            {
                var sets = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, T>>(_user.Data[_storeIdentifier].Value);
                if (sets != null)
                {
                    values = sets;
                }
            }

            if (values.ContainsKey(key) == false)
            {
                values.Add(key, value);
            }
            else 
            {
                values[key] = value;
            }

            _user.Data.Apply(_storeIdentifier, System.Text.Json.JsonSerializer.Serialize(values));

            await _user.SaveAsync();
        }
    }
}
