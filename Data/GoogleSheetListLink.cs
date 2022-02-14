using System.Data;
using Sushi.MicroORM;
using Sushi.MicroORM.Mapping;

namespace Sushi.Mediakiwi.Module.GoogleSheetsSync.Data
{
    [DataMap(typeof(GoogleSheetListLinkMap))]
    public class GoogleSheetListLink
    {
        public class GoogleSheetListLinkMap : DataMap<GoogleSheetListLink>
        {
            public GoogleSheetListLinkMap()
            {
                Table("cat_GoogleSheetListLinks");
                Id(x => x.ID, "GoogleSheetListLink_Key").Identity();
                Map(x => x.UserID, "GoogleSheetListLink_User_Key");
                Map(x => x.ListID, "GoogleSheetListLink_List_Key");
                Map(x => x.SheetId, "GoogleSheetListLink_Sheet_Id").Length(50);
                Map(x => x.SheetUrl, "GoogleSheetListLink_Sheet_Url").Length(512);
                Map(x => x.LastExport, "GoogleSheetListLink_LastExport");
                Map(x => x.LastImport, "GoogleSheetListLink_LastImport");
            }
        }

        public int ID { get; set; }
        public int UserID { get; set; }
        public int ListID { get; set; }
        public string SheetId { get; set; }
        public string SheetUrl { get; set; }
        public DateTime? LastExport { get; set; }
        public DateTime? LastImport { get; set; }


        public static List<GoogleSheetListLink> FetchAll()
        {
            var connector = new Connector<GoogleSheetListLink>();
            var filter = connector.CreateQuery();
            var result = connector.FetchAll(filter);
            return result;
        }

        public static async Task<List<GoogleSheetListLink>> FetchAllAsync()
        {
            var connector = new Connector<GoogleSheetListLink>();
            var filter = connector.CreateQuery();
            var result = await connector.FetchAllAsync(filter);
            return result;
        }

        public static GoogleSheetListLink FetchSingle(int id)
        {
            var connector = new Connector<GoogleSheetListLink>();
            var result = connector.FetchSingle(id);
            return result;
        }

        public static async Task<GoogleSheetListLink> FetchSingleAsync(int id)
        {
            var connector = new Connector<GoogleSheetListLink>();
            var result = await connector.FetchSingleAsync(id);
            return result;
        }

        public static async Task<GoogleSheetListLink> FetchSingleAsync(int listId, int userId)
        {
            var connector = new Connector<GoogleSheetListLink>();
            var filter = connector.CreateQuery();
            filter.Add(x => x.ListID, listId);
            filter.Add(x => x.UserID, userId);

            var result = await connector.FetchSingleAsync(filter);
            return result;
        }

        public void Save()
        {
            var connector = new Connector<GoogleSheetListLink>();
            connector.Save(this);
        }

        public async Task SaveAsync()
        {
            var connector = new Connector<GoogleSheetListLink>();
            await connector.SaveAsync(this);
        }

        public void Delete()
        {
            var connector = new Connector<GoogleSheetListLink>();
            connector.Delete(this);
        }

        public async Task DeleteAsync()
        {
            var connector = new Connector<GoogleSheetListLink>();
            await connector.DeleteAsync(this);
        }
    }
}