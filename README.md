# Simple python program to get data from an Allen Bradley PLC to an excel sheet.

I have added the small test program( v32) and the excel sheet it creates.

Revised the GUI for much cleaner look.

4-11-25 removed unused import from main file.
4-20-25:  Added Gridview like table at bottom of page to keep track in real time the data being read.

[ALLEN BRADLEY PLC TO EXCEL INSTRUCTIONS.docx](https://github.com/user-attachments/files/19827565/ALLEN.BRADLEY.PLC.TO.EXCEL.INSTRUCTIONS.docx)


NEW!!!!!!!!!
PLCdata_da.py file.

This file now contains some data analysis on a separate gui page. you can load an excel sheet, pick a column from excel sheet to get sum, avg, or cpk from that column of data.

FINAL UPDATE 6-13-25!!! I have added a checkbox to create a new excel sheet at midnight so you can continuously gather data and keep it separated by day.
Also added a graph feature on data page, you can select starting from the beginning of the file how many points you want to plot.
This will likely be the final update for this program as its gotten bigger than I expected and really should have been put into classes to make 
it more readable.  Oh well.  New pics of updates at end.


FIRST ! you must change filepath to where you want data stored.


![changepath](https://github.com/user-attachments/assets/eb0901c5-c0b9-4b61-841f-f1b5ef603e09)




Set IP of your PLC, hit Save IP


![BeginIP](https://github.com/user-attachments/assets/0d2db199-a17c-4854-9d1a-d3ed70608554)





IF valid IP address it will say saved


![IPsaved](https://github.com/user-attachments/assets/c4da328f-4373-4462-92c8-95172b036cc7)








Add your collection interval in seconds.


![addintervalsec](https://github.com/user-attachments/assets/e48c3ce1-87cc-4d30-83ce-93706e418e37)










Hit get tags to get all tags from PLC.


![Getalltags](https://github.com/user-attachments/assets/02e6096b-a25d-495d-9ae5-0d5f5c028c73)











Add your column names ( up to 8 ).  If you dont add them they will be blank!! Hit change to change name.
If ok the name textbox should go blank and you select your next column name.


![column_name_Change](https://github.com/user-attachments/assets/be0a9ab9-71d8-465f-b947-534b4d00feb6)











Add what you data you want to retrieve--be it an array, single tag, udt, ot multi-single tags.
The array sintax is arrayname{n} n=number of indexes to read.  udt is just udt name, as well as single
and multitags are just name.  Hit button next to Entry box to start collecting data.


![AddArraytoget](https://github.com/user-attachments/assets/16a0a79f-a4be-4939-8a95-d3860080e2ff)









You should see data in the data(tree view ) now.


![dataview](https://github.com/user-attachments/assets/c3a73e01-dbfd-4905-82e1-3d132d0d0104)










Hit stop logging button to stop logging.


![stoplogpython](https://github.com/user-attachments/assets/74835f5c-ffa8-4590-b286-6abf5a497976)












To analyze data hit the data button.


![databtn](https://github.com/user-attachments/assets/14495334-c729-481d-84de-d41d2727a803)










select excel file you want to analyze. You should be able to open any excel file not just ones made with this program.


![dataopenxl](https://github.com/user-attachments/assets/64e2bb0d-17b5-46f5-9f01-e4e7cb3bae18)











The headers of the columns should show up ( if they are the top row!! )


![headersdata](https://github.com/user-attachments/assets/901c7dd8-096d-4fe2-9bc2-783985102872)










Enter the column you want to analyze and hit get column the textbox should load with data.


![columndata](https://github.com/user-attachments/assets/5b7f17dd-7bfb-4050-8308-96ebc4ce72a4)










Hit Average button to get the average of the column.


![avgdata](https://github.com/user-attachments/assets/5fe1cc48-d017-4715-b1cf-3e4440f16a70)









Hit the sum button to get some of the column


![sumdata](https://github.com/user-attachments/assets/62a90735-26f8-4f53-9f66-a47980244e90)










Enter Lower spec Limit and Upper spec limit for CPK, then hit cpk to get cpk of column. Just
dont look at my numbers, theyre backwords for usl and lsl!  Hit close to close the data gui.

![cpkData](https://github.com/user-attachments/assets/20611650-cdfd-4608-94a8-d1a2a1e9019e)









UPDATE 6-13-25.
Check this box if you want to create new sheets at midnight.

![createnewsheet](https://github.com/user-attachments/assets/1b01d861-ec25-4b67-9d88-9310d2f808cb)








You can now select graph to graph out points you want from the column data you selected.


![graph](https://github.com/user-attachments/assets/e457398d-7b88-4ae7-b527-318d8eca5fec)









