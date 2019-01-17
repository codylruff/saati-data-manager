SELECT * FROM tblWarpingSpecs WHERE MaterialNumber = OKE101011AP00CN;

SELECT * FROM tblStyleSpecs 
WHERE Style = 10;

INSERT INTO tblWarpingSpecs (MaterialNumber, MaterialDescription, FinalWidthCm, NumberOfEnds, IsSWrapped, SpringColor, WarpingSpeed, BeamingSpeed, CrossWinding, DentsPerCm, EndsPerDent, Style, BeamWidth, YarnSupplier, YarnCode, K1, WarpingTension, K2, BeamingTension, Time_Stamp) 
VALUES ('OKE101011AP00CN', 'STY 101 K29 3300D 1F279 131CM', '131', '2227', 'True', 'Yellow', '300', '80', '10', '0', '0', '101', '130.2', 'Dupont', 'FKE002933005NP0000', '0.25', '0', '1.25', '0', '1/16/2019 3:46:24 PM')