--create database IndicatorBook

use IndicatorBook
CREATE TABLE Sent (
    Id int IDENTITY(1,1) PRIMARY KEY NOT NULL,
    [File] nvarchar(50) NOT NULL,
    Pamphleteer nvarchar(50) NOT NULL,
    Number int NOT NULL,
    Title nvarchar(255) NOT NULL,
	Description nvarchar(255),
	NextNumber int not null,
    HasAttachment BIT,
    [Date] nvarchar(255) not null,
	WordFile VARBINARY(MAX) not null,
    WordFileExtension nvarchar(50) not null
);


create unique index Sent_File_Pamphleteer_Number
on Sent([File], Pamphleteer, Number);