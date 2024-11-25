alter table pargen add PAR_IGV smallint

UPDATE PARGEN SET PAR_IGV = 10 where PAR_CODCIA = '01'
UPDATE PARGEN SET PAR_IGV = 10 where PAR_CODCIA = '02'