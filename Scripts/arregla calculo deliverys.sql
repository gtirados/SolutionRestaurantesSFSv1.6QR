select FAR_SUBTOTAL, FAR_IMPTO,FAR_BRUTO, * from FACART where FAR_FBG = 'F' AND FAR_NUMFAC = 3

UPDATE FACART SET FAR_BRUTO = ROUND(FAR_SUBTOTAL/1.10,2) where FAR_FBG = 'F' AND FAR_NUMFAC = 3
UPDATE FACART SET FAR_IMPTO = ROUND(FAR_BRUTO*0.10,2) where FAR_FBG = 'F' AND FAR_NUMFAC = 3