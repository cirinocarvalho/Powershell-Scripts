-- Target table for examples/Sync-EarthquakeData.ps1
IF OBJECT_ID(N'dbo.Earthquake', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.Earthquake
    (
        EventId      NVARCHAR(64)   NOT NULL PRIMARY KEY,
        Magnitude    DECIMAL(4, 2)  NULL,
        Place        NVARCHAR(256)  NULL,
        EventTimeUtc DATETIME2(3)   NULL,
        Longitude    DECIMAL(9, 5)  NULL,
        Latitude     DECIMAL(9, 5)  NULL,
        DepthKm      DECIMAL(8, 3)  NULL,
        DetailUrl    NVARCHAR(512)  NULL,
        LoadedUtc    DATETIME2(3)   NOT NULL CONSTRAINT DF_Earthquake_LoadedUtc DEFAULT SYSUTCDATETIME()
    );

    CREATE INDEX IX_Earthquake_EventTimeUtc ON dbo.Earthquake (EventTimeUtc DESC);
END
GO
