			var Upl_IOfun,Upl_Level=0,Upl_GetDelay=3000,Upl_Start;			
			var Upl_id=1,Upl_url = "";
			
			var Upl_IOfun,Upl_Level=0,Upl_GetDelay=3000,Upl_Start;			
			
			function GetTimeString(Num)
			{
				var Str="",Temp,Number;
				Number = Num;
				Temp = Number/(24*60*60);
				if(parseInt(Temp) > 0)
				Str = Str + parseInt(Temp) + "��";
				
				Number = Number-parseInt(Temp)*24*60*60;
				Temp = Number/(60*60);
				if(parseInt(Temp) > 0)
				Str = Str + parseInt(Temp) + "ʱ";
			
				Number = Number-parseInt(Temp)*60*60;
				Temp = Number/(60);
				if(parseInt(Temp) > 0)
				Str = Str + parseInt(Temp) + "��";
			
				Number = Number-parseInt(Temp)*60
				Temp = parseInt(Number)
				if(parseInt(Temp) > 0)
				Str = Str + Temp + "��";
				if(Str == "")
				Str = "������...";
				return Str;
			
			}
			function Upl_IO(ur,lb,id)
			{
				Upl_Level += 1;
				getAJAX(Upl_url+"&tt=" + Math.random(),"","Upl_IO_processor(tmp);",1);
				window.clearTimeout(Upl_IOfun);
				if(Upl_Level<2)Upl_IOfun = window.setTimeout(Upl_IO,Upl_GetDelay);
				Upl_Level -= 1;
			}
			function Upl_IO_processor(str)
			{
				if(str==" "){window.clearTimeout(Upl_IOfun);Upl_done();return;}

				var tp="upload_doc";
				var tmp;
				if(str!=" " && str != "none")
				{
					Upl_Start = true;
					
					if(str!="start")
					{
						tmp = str.split("|");
						if(tmp.length>=5)
						{
							$id('img'+Upl_id).width=tmp[0];
							$id('txt'+Upl_id).innerHTML=tmp[1];
							$id('tm'+Upl_id).innerHTML="����ʱ:" + GetTimeString(tmp[2]) + " ����ʣ��:" + GetTimeString(tmp[3]);
							$id('img'+Upl_id).title="(" + tmp[4] + ")";
						}
						if(tmp.length>=6)$id('bartitle'+Upl_id).innerHTML="" + tmp[5] + "";
					}
				}
				else
				{
					if(Upl_Start)
					{
						$id('tm'+Upl_id).innerHTML = "��ǰ������ɣ����Ժ�...";
						Upl_Level=9999;window.clearTimeout(Upl_IOfun);
					}
					else
					{
						$id('tm'+Upl_id).innerHTML = "��ǰ��������У����Ժ�...";
					}
				}
			}
			function Upl_done()
			{
				window.clearTimeout(Upl_IOfun);
				getAJAX(Upl_url+"&free=1&tt=" + Math.random(),"","",1);
				
				$id('img'+Upl_id).width=400;
				$id('txt'+Upl_id).innerHTML="100.00";
				$id('tm'+Upl_id).innerHTML="�������.";
			}
			Upl_Start = false;Upl_Level=0;