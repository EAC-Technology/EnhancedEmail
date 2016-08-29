/*==========================
dialog_v2_callback JS functions
===========================*/


jQuery(document).ready(function() {
	 jQuery(e2vdomSV['DialogGUID']).dialog('registerCallbacks', {
			'selectEltOfPath' : function(data){
							eval('data='+data);
							execEventBinded('8fa63233_68c8_43ea_84fe_d36a257be1fb', 'itemclick', data);
			},
			'runProgress' : function(){						
						$topLoader.setProgress(data);						
			}
		})
	}
)


function login (token) {
	apidata ={};

	$.post( "restapi.py", { 
							appid: "b0a274f0-22bc-44be-be48-da6ec9180268", 
							objid: "5073ff75-da99-44fb-a5d7-e44e5ab28598", 
							action_name: "login",
							xml_data : "{\"login\": \"root\", \"password\": \"root\"}", 
							callback : "parseReturn" })
		.done(function( data ) {
		   eval(data);
        });
}

function getAnswer(objanswer) {
	if (objanswer[0] == "success") {
		return objanswer[1];
	}
}

function showProgress() {
	$('#topLoader').css('visibility','visible');
	topLoaderRunning = true;
}

function hideProgress() {
	$('#topLoader').css('visibility','hidden');
	topLoaderRunning = false;
}

function getProgress(ProgressID) {
	apidata ={};
	apidata["plugin_guid"] = "95533a74-bbe3-4962-9405-9aa96a4092cd";
	apidata["async"]=false;
	apidata["name"]="getprogress";
	apidata["data"]=ProgressID;
	
	cnxdata = { 
				appid: "b0a274f0-22bc-44be-be48-da6ec9180268", 
				objid: "5073ff75-da99-44fb-a5d7-e44e5ab28598", 
				action_name: "call_macro",
				xml_data : JSON.stringify(apidata), 
				callback : "getAnswer" 
			};
	
	$.post( "restapi.py", cnxdata )
		.done(function( data ) {
			progressData = JSON.parse(eval(data.replace("/**/","")));
			$topLoader.setValue(progressData["label"]);
		    $topLoader.setProgress(progressData["value"]);
			if($('.ui-widget-overlay').css('z-index'))
				{if ($('#topLoader').css('visibility')=='visible')
					{t = setTimeout(function(){getProgress(ProgressID)}, 500);}
				} 
        }); 
}

function getTable(csvname,pos,token) {
	apidata ={};
	apidata["plugin_guid"] = "95533a74-bbe3-4962-9405-9aa96a4092cd";
	apidata["async"]=false;
	apidata["name"]="gettable";
	funcdata ={};
	funcdata["csvname"] = csvname;
	funcdata["pos"] = pos;
	funcdata["token"] = token;
	apidata["data"]=JSON.stringify(funcdata);

	cnxdata = { 
				appid: "b0a274f0-22bc-44be-be48-da6ec9180268", 
				objid: "5073ff75-da99-44fb-a5d7-e44e5ab28598", 
				action_name: "call_macro",
				xml_data : JSON.stringify(apidata), 
				callback : "getAnswer" 
			};
			
	$.post( "restapi.py", cnxdata )
		.done(function( data ) {
			htmlTable = JSON.parse(eval(data.replace("/**/","")));
			$("#datagrid_"+token).html(htmlTable["code"])
		});	
	
}