use IndicatorBook
CREATE TABLE Recived (
    Id int IDENTITY(1,1) PRIMARY KEY NOT NULL,
    RowNumber int NOT NULL,
    PreviousRowNumber int NOT NULL,
    [Date] nvarchar(255) not null,
    LetterOwners nvarchar(255) NOT NULL,
	Description nvarchar(255),
    HasAttachment BIT,
	RecivedLetterNumber nvarchar(255) NOT NULL,
    RecivedLetterDate nvarchar(255) not null,
	ScanFile VARBINARY(MAX) not null,
    ScanFileExtension nvarchar(50) not null
);


create unique index Recived_RowNumber
on Recived(RowNumber);