CREATE TABLE [dbo].[cat_GoogleSheetListLinks](
	[GoogleSheetListLink_Key] [int] IDENTITY(1,1) NOT NULL,
	[GoogleSheetListLink_User_Key] [int] NOT NULL,
	[GoogleSheetListLink_List_Key] [int] NOT NULL,
	[GoogleSheetListLink_Sheet_Id] [nvarchar](50) NULL,
	[GoogleSheetListLink_Sheet_Url] [nvarchar](512) NULL,
	[GoogleSheetListLink_LastExport] [datetime] NULL,
	[GoogleSheetListLink_LastImport] [datetime] NULL,
 CONSTRAINT [PK_cat_GoogleSheetListLinks] PRIMARY KEY CLUSTERED 
(
	[GoogleSheetListLink_Key] ASC
)WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY];


ALTER TABLE [dbo].[cat_GoogleSheetListLinks]  WITH CHECK ADD  CONSTRAINT [FK_cat_GoogleSheetListLinks_wim_ComponentLists] FOREIGN KEY([GoogleSheetListLink_List_Key])
REFERENCES [dbo].[wim_ComponentLists] ([ComponentList_Key]);


ALTER TABLE [dbo].[cat_GoogleSheetListLinks] CHECK CONSTRAINT [FK_cat_GoogleSheetListLinks_wim_ComponentLists];


ALTER TABLE [dbo].[cat_GoogleSheetListLinks]  WITH CHECK ADD  CONSTRAINT [FK_cat_GoogleSheetListLinks_wim_Users] FOREIGN KEY([GoogleSheetListLink_User_Key])
REFERENCES [dbo].[wim_Users] ([User_Key]);


ALTER TABLE [dbo].[cat_GoogleSheetListLinks] CHECK CONSTRAINT [FK_cat_GoogleSheetListLinks_wim_Users];
