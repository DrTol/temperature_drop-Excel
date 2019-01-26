# Temperature Drop through Insulated Pipe (Steady State)
This VBA function/Excel tool calculates the temperature drop through a pre-insulated buried single steel pipe (steady-state). 

## Table of Contents
- [How2Use](README.md#how2use)
- [License](README.md#license)
- [Acknowledgement](README.md#acknowledgement)
- [How2Cite](README.md#how2cite)
- [References](README.md#references)

## How2Use
An example Excel file is given: [TemperatureDrop_Buried&InsulatedPipeV02_r02.xlsm](https://github.com/DrTol/temperature_drop-Excel/blob/master/TemperatureDrop_Buried%26InsulatedPipeV02_r02.xlsm), which illustrates how to use the VB functions developed i) [Tout](https://github.com/DrTol/temperature_drop-Excel/blob/master/Tout) and ii) [Tout_simple](https://github.com/DrTol/temperature_drop-Excel/blob/master/Tout_simple). 

These functions are based on other VBA functions: 
a) [XSteam](https://github.com/DrTol/temperature_drop-Excel/blob/master/XSteam) for the water properties, 
b) [f_Clamond](https://github.com/DrTol/temperature_drop-Excel/blob/master/f_Clamond) to return the Darcy friction factor,
c) [zOthers](https://github.com/DrTol/temperature_drop-Excel/blob/master/zOthers) involves of VB functions to calculate pi number and Reynolds number

If you are new to Visual Basic (Excel), please follow the instructions in [Copy a macro module to another workbook](https://support.office.com/en-us/article/copy-a-macro-module-to-another-workbook-13c0938b-8432-4259-9177-a71f7e626de0) OR simply make use of the example Excel file [TemperatureDrop_Buried&InsulatedPipeV02_r02.xlsm](https://github.com/DrTol/temperature_drop-Excel/blob/master/TemperatureDrop_Buried%26InsulatedPipeV02_r02.xlsm) to base on for your calculations. 

## License
You are free to use, modify and distribute the code as long as the authorship is properly acknowledged. The same applies for the original works 'XSteam' by Holmgren M and 'f_Clamond' by Clamond, D. 

## Acknowledgement 
In memory of my father Bekir Tol..

We would like to acknowledge all of the open-source minds in general for their willing of share (as apps or comments/answers in forums), which has encouraged our department to publish our tools developed.

## How2Cite:
1. Tol, Hİ. temperature_drop-Excel. DOI: XXX. GitHub Repository 2018; https://github.com/DrTol/temperature_drop-Excel

## References
- Cengel, YA. Heat transfer: A practical approach (2. ed)
- Kwon KC, 1998. Heat transfer model of above and underground insulated piping systems. In: ASME conference - Heat exchanger committee meeting & International joint power generation conference, Baltimore, MA. 
- LOGSTOR, LOGSTOR calculator - Energy Loss, Available at: http://calc.logstor.com/
- Ankur, Heat transfer in buried liquid pipelines, Available at: https://www.cheresources.com/invision/blog/4/entry-493-heat-transfer-in-buried-liquid-pipelines/
- Breizh, insulated pipe buried V1.xlsx, Available at: https://www.cheresources.com/invision/index.php?app=core&module=attach&section=attach&attach_id=12301
- Kusuda T, 1981. Heat transfer analysis of underground heat and chilled-water distribution systems. Report No: NBSIR 81-2378, US Department of Commerce, Washington DC.
- Subramanian RS, Heat transfer in flow through conduits, Clarkson University. 
