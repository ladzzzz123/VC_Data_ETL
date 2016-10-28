# VC_Research
Venture capital data ETL ：  python &amp; sqlite

Member of research team lead by associate professor Q.Li from Zhejiang University City College.

Data Description:

Raw Data is downloaded from Econ. databases, semi-structured data containing information chunks including: 
1. VC organization description: Name, Type, Found time, HQ position, Website etc..
2. Investment information: target enterprises, time, series, industry, district etc..
3. Cooperation information.
4. VC Exit information: Type (s.a. IPO, Mergers and Acquisition), time, etc..

One orgnization per excel sheet, and there's 6 XLS files averaging 500 sheets / file, designed 2 loops to read each sheet as a "sparse" dataframe. Transformed the semi structured chunks into structured T1-T4 for further use.

Further operation:
