<img src="https://user-images.githubusercontent.com/45975234/235141428-91ee5bfb-5b94-4f8d-a2db-a92a0f024d25.png" height="100" >

# RFNSA

A Python tool to wrangle antenna details from RFNSA STAD Table for easier and faster populating into EME compliance software and EMEG Equipment List with accuracy and transparency saving you loads of work and massive amount of time.


## Example Input & Output

### Input
Click the "Download all STADs" on STAD tab of your site on RFNSA and you'll get a compressed txt file. Press ctrl + A and ctrl + C in text file, press ctrl + V in an excel sheet and save it. Replace excel sheet name and path in this python script. Run All

### Output

The output will be an excel file with the following information in separate sheets. 

#### Add following antennas to Prox5 and label them
> you might want to set powers/frequencies to ports of each single antenna type before labelling them, so you can clone similar ones and then label it. 

![image](https://user-images.githubusercontent.com/45975234/236609858-b07c11fa-7b7d-427a-a1ea-a542447a112b.png)


#### Set height, bearing, and MDT to antennas 

![image](https://user-images.githubusercontent.com/45975234/236609872-f967d477-1382-4894-8921-1d124c2b57cc.png)


#### Set powers and frequencies to antennas 

![image](https://user-images.githubusercontent.com/45975234/236652349-30b22881-5d79-4cdb-bb6e-208da8e9b445.png)


#### EMEG "Equipment List": just copy following technologies/sector and power (W) of all antennas

![image](https://user-images.githubusercontent.com/45975234/236651854-aa1524db-d6a5-43a8-b650-6d133ac25dda.png)

#### License

The source code and dataset are available under Creative Commons BY-NC 4.0 license by NAVER Corporation. You can use, copy, tranform and build upon the material for non-commercial purposes as long as you give appropriate credit by citing our paper, and indicate if changes were made.

For technical and business inquiries, please contact hharis11@hotmail.com.

