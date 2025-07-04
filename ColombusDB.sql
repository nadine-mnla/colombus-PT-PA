CREATE TABLE Kamar (
    ID INT PRIMARY KEY,
    NoKamar VARCHAR(10) UNIQUE,
    Tipe VARCHAR(20) NOT NULL,
    Status VARCHAR(15) CHECK (Status IN ('Tersedia', 'Terisi', 'Maintenance'))
);

CREATE TABLE Tamu (
    ID INT PRIMARY KEY,
    Nama VARCHAR(50) NOT NULL,
    NoKTP VARCHAR(20),
    KamarID INT FOREIGN KEY REFERENCES Kamar(ID)
);