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

![image](https://github.com/hassanharis/RFNSA-STAD-Wrangler/assets/45975234/830267c8-84f9-4093-a639-5ed3aca63910)

#### Set powers and frequencies to antennas 

![image](https://github.com/hassanharis/RFNSA-STAD-Wrangler/assets/45975234/5ec8a97c-313f-44d8-9acb-b3cdb905a50e)

#### Set bearing, height, and MDT to antennas 

![image](https://github.com/hassanharis/RFNSA-STAD-Wrangler/assets/45975234/8851cc88-70bb-4a8e-bd3f-0c019042dd1e)


#### EMEG "Equipment List": just copy following technologies/sector and power (W) of all antennas

![image](https://github.com/hassanharis/RFNSA-STAD-Wrangler/assets/45975234/4bce12d0-04c2-470a-81b0-e1dd1d995cf0)

#### CANRAD: validate this CANRAD version of RFNSA with Telstra provided CANRAD 

![image](https://github.com/hassanharis/RFNSA-STAD-Wrangler/assets/45975234/06dea5b6-b776-453a-a8f5-cffeef8c6830)


#### License

The source code and dataset are available under Creative Commons BY-NC 4.0 license by NAVER Corporation. You can use, copy, tranform and build upon the material for **non-commercial purposes** as long as you give appropriate credit by citing our paper, and indicate if changes were made.

For technical and business inquiries, please contact hharis11@hotmail.com.

