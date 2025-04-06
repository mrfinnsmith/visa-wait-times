function processVisaWaitTimes() {
 var MAX_RUNTIME = 4 * 60 * 1000; // 4 minutes max runtime
 var START_TIME = new Date().getTime();
 var properties = PropertiesService.getScriptProperties();
 var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
 
 // Check if we've completed today's run
 var lastRunDate = properties.getProperty('lastCompletedDate');
 if (lastRunDate === today) {
   return; // Already completed today's run
 }
 
 var postData = {
   'Abidjan': 'abidjan',
   'Abu Dhabi': 'P2',
   'Abuja': 'P3',
   'Accra': 'P4',
   'Adana': 'adana',
   'Addis Ababa': 'P5',
   'Algiers': 'P6',
   'Almaty': 'P7',
   'Amman': 'P8',
   'Amsterdam': 'P9',
   'Ankara': 'P10',
   'Antananarivo': 'P11',
   'Apia': 'P225',
   'Ashgabat': 'P12',
   'Asmara': 'P13',
   'Astana': 'astana',
   'Asuncion': 'P15',
   'Athens': 'athens',
   'Auckland': 'P17',
   'Baghdad': 'P226',
   'Baku': 'P19',
   'Bamako': 'P20',
   'Bandar Seri Begawan': 'P21',
   'Bangkok': 'P22',
   'Bangui': 'bangui',
   'Banjul': 'banjul',
   'Barcelona': 'barcelona',
   'Beijing': 'P24',
   'Beirut': 'P25',
   'Belfast': 'P26',
   'Belgrade': 'P27',
   'Belmopan': 'P28',
   'Berlin': 'P29',
   'Bern': 'P30',
   'Bishkek': 'P31',
   'Bogota': 'P32',
   'Brasilia': 'P33',
   'Bratislava': 'P34',
   'Brazzaville': 'P35',
   'Bridgetown': 'P36',
   'Brussels': 'P37',
   'Bucharest': 'P38',
   'Budapest': 'P39',
   'Buenos Aires': 'P40',
   'Bujumbura': 'P41',
   'Cairo': 'P42',
   'Calgary': 'P43',
   'Canberra': 'canberra',
   'Cape Town': 'P44',
   'Caracas': 'P45',
   'Casablanca': 'P46',
   'Chengdu': 'P47',
   'Chennai (Madras)': 'P48',
   'Chiang Mai': 'P49',
   'Chisinau': 'P50',
   'Ciudad Juarez': 'P51',
   'Colombo': 'P52',
   'Conakry': 'P53',
   'Copenhagen': 'P54',
   'Cotonou': 'P55',
   'Curacao': 'P223',
   'Dakar': 'P56',
   'Damascus': 'P57',
   'Dar Es Salaam': 'P58',
   'Dhahran': 'P59',
   'Dhaka': 'P60',
   'Dili': 'P227',
   'Djibouti': 'P61',
   'Doha': 'P62',
   'Dubai': 'P63',
   'Dublin': 'P64',
   'Durban': 'P65',
   'Dushanbe': 'P66',
   'Edinburgh': 'edinburgh',
   'Erbil': 'erbil',
   'Florence': 'P67',
   'Frankfurt': 'P68',
   'Freetown': 'P69',
   'Fukuoka': 'fukuoka',
   'Gaborone': 'P70',
   'Georgetown': 'P71',
   'Guadalajara': 'P72',
   'Guangzhou': 'P73',
   'Guatemala City': 'P74',
   'Guayaquil': 'P75',
   'Halifax': 'P76',
   'Hamilton': 'P77',
   'Hanoi': 'P78',
   'Harare': 'P79',
   'Havana': 'P80',
   'Helsinki': 'P81',
   'Hermosillo': 'P82',
   'Ho Chi Minh City': 'P83',
   'Hong Kong': 'P84',
   'Hyderabad': 'P85',
   'Islamabad': 'P86',
   'Istanbul': 'P87',
   'Jakarta': 'P88',
   'Jeddah': 'P89',
   'Jerusalem': 'P90',
   'Johannesburg': 'P91',
   'Juba': 'P228',
   'Kabul': 'P229',
   'Kampala': 'P93',
   'Kaohsiung': 'kaohsiung',
   'Karachi': 'P94',
   'Kathmandu': 'P95',
   'Khartoum': 'P96',
   'Kigali': 'P97',
   'Kingston': 'P98',
   'Kinshasa': 'P99',
   'Kolkata': 'P100',
   'Kolonia': 'P101',
   'Koror': 'P102',
   'Krakow': 'P103',
   'Kuala Lumpur': 'P104',
   'Kuwait': 'P105',
   'Kyiv': 'P106',
   'La Paz': 'P107',
   'Lagos': 'P108',
   'Lahore': 'lahore',
   'Libreville': 'P109',
   'Lilongwe': 'P110',
   'Lima': 'P111',
   'Lisbon': 'P112',
   'Ljubljana': 'P113',
   'Lome': 'lome',
   'London': 'P115',
   'Luanda': 'P116',
   'Lusaka': 'P117',
   'Luxembourg': 'P118',
   'Madrid': 'P119',
   'Majuro': 'P120',
   'Malabo': 'P121',
   'Managua': 'managua',
   'Manama': 'P123',
   'Manila': 'P124',
   'Maputo': 'P125',
   'Marseille': 'marseille',
   'Maseru': 'P126',
   'Matamoros': 'P127',
   'Mbabane': 'P128',
   'Melbourne': 'P129',
   'Merida': 'P130',
   'Mexicali Tpf': 'mexicali_tpf',
   'Mexico City': 'P131',
   'Milan': 'P132',
   'Minsk': 'P133',
   'Monrovia': 'P134',
   'Monterrey': 'P135',
   'Montevideo': 'P136',
   'Montreal': 'P137',
   'Moscow': 'P138',
   'Mumbai (Bombay)': 'P139',
   'Munich': 'P140',
   'Muscat': 'P141',
   'N\'Djamena': 'P142',
   'Naha': 'P143',
   'Nairobi': 'P144',
   'Naples': 'P145',
   'Nassau': 'P146',
   'New Delhi': 'P147',
   'Niamey': 'P148',
   'Nicosia': 'P149',
   'Nogales': 'P150',
   'Nouakchott': 'P151',
   'Nuevo Laredo': 'P152',
   'Osaka/Kobe': 'P153',
   'Oslo': 'P154',
   'Ottawa': 'P155',
   'Ouagadougou': 'P156',
   'Panama City': 'P157',
   'Paramaribo': 'P158',
   'Paris': 'P159',
   'Perth': 'P160',
   'Phnom Penh': 'P161',
   'Podgorica': 'P162',
   'Ponta Delgada': 'ponta_delgada',
   'Port Au Prince': 'P164',
   'Port Louis': 'P165',
   'Port Moresby': 'P166',
   'Port Of Spain': 'P167',
   'Porto Alegre': 'porto_alegre',
   'Prague': 'P168',
   'Praia': 'P169',
   'Pretoria': 'pretoria',
   'Pristina': 'P231',
   'Quebec': 'P170',
   'Quito': 'P171',
   'Rangoon': 'P172',
   'Recife': 'P173',
   'Reykjavik': 'P174',
   'Riga': 'P175',
   'Rio De Janeiro': 'P230',
   'Riyadh': 'P177',
   'Rome': 'P178',
   'San Jose': 'P179',
   'San Salvador': 'P180',
   'Sanaa': 'P181',
   'Santiago': 'P182',
   'Santo Domingo': 'P183',
   'Sao Paulo': 'P184',
   'Sapporo': 'P224',
   'Sarajevo': 'P185',
   'Seoul': 'P186',
   'Shanghai': 'P187',
   'Shenyang': 'P188',
   'Singapore': 'P189',
   'Skopje': 'P190',
   'Sofia': 'P191',
   'St Georges': 'st_georges',
   'St Petersburg': 'P192',
   'Stockholm': 'P193',
   'Surabaya': 'P194',
   'Suva': 'P195',
   'Sydney': 'P196',
   'Taipei': 'P197',
   'Tallinn': 'P198',
   'Tashkent': 'P199',
   'Tbilisi': 'P200',
   'Tegucigalpa': 'P201',
   'Tel Aviv': 'P202',
   'Tijuana': 'tijuana',
   'Tijuana Tpf': 'P203',
   'Tirana': 'P204',
   'Tokyo': 'P205',
   'Toronto': 'P206',
   'Tripoli': 'P207',
   'Tunis': 'P208',
   'Ulaanbaatar': 'P209',
   'Usun-New York': 'usun-new_york',
   'Valletta': 'P210',
   'Vancouver': 'P211',
   'Vienna': 'P212',
   'Vientiane': 'P213',
   'Vilnius': 'P214',
   'Vladivostok': 'P215',
   'Warsaw': 'P216',
   'Washington Refugee Processing Center': 'washington_refugeeprocessingcenter',
   'Windhoek': 'P217',
   'Wuhan': 'wuhan',
   'Yaounde': 'P218',
   'Yekaterinburg': 'yekaterinburg',
   'Yerevan': 'P220',
   'Zagreb': 'P221'
 };
 
 var posts = Object.keys(postData);
 var batchSize = 20;
 
 // Get current position
 var currentIndex = parseInt(properties.getProperty('currentIndex') || 0);
 var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 var sheet = spreadsheet.getSheetByName('API');
 
 // Create sheet with headers if needed
 if (!sheet) {
   sheet = spreadsheet.insertSheet('API');
   var headers = [
     'Post',
     'CID',
     'B1/B2 Visitor',
     'F/M/J Student',
     'Other',
     'C/D Crew',
     'B1/B2 Waiver',
     'Column6',
     'Column7',
     'Column8',
     'Date'
   ];
   sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
 }
 
 // Find the last row with data
 var lastRow = Math.max(1, sheet.getLastRow());
 
 // Start adding new data after the last row
 var row = lastRow + 1;
 
 // Process each post in this batch
 var endIndex = Math.min(currentIndex + batchSize, posts.length);
 for (var i = currentIndex; i < endIndex; i++) {
   // Check if we're approaching max runtime
   if (new Date().getTime() - START_TIME > MAX_RUNTIME) {
     properties.setProperty('currentIndex', i.toString());
     return; // Exit to avoid timeout
   }
   
   var post = posts[i];
   try {
     var cid = postData[post];
     var url = "https://travel.state.gov/content/travel/resources/database/database.getVisaWaitTimes.html?cid=" + cid;
     var response = UrlFetchApp.fetch(url);
     var data = response.getContentText();
     
     if (data && !data.toLowerCase().includes("not found")) {
       // Split the response by pipe and clean up the numbers
       var waitTimes = data.split('|').map(function(time) {
         // Extract just the number and trim
         var match = time.match(/(\d+)/);
         return match ? match[1] : '';
       });
       
       // Create row with post, cid, wait times, and date
       var rowData = [post, cid].concat(waitTimes).concat([today]);
       sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
       row++;
     }
     
     // Add small delay to avoid rate limits
     Utilities.sleep(100);
     
   } catch (error) {
     Logger.log("Error processing " + post + " (" + cid + "): " + error.toString());
   }
 }
 
 // Save current position or mark as complete
 if (endIndex >= posts.length) {
   // All done for today
   properties.setProperty('lastCompletedDate', today);
   properties.deleteProperty('currentIndex');
   sheet.autoResizeColumns(1, 11); // Resize columns to fit the data
 } else {
   properties.setProperty('currentIndex', endIndex.toString());
 }
}

function resetVisaDataProgress() {
 var properties = PropertiesService.getScriptProperties();
 properties.deleteProperty('currentIndex');
 properties.deleteProperty('lastCompletedDate');
}

function testVisaAPI() {
 var cid = "P24"; // Beijing as test case
 var url = "https://travel.state.gov/content/travel/resources/database/database.getVisaWaitTimes.html?cid=" + cid;
 var response = UrlFetchApp.fetch(url);
 var data = response.getContentText();
 Logger.log("Raw API response: " + data);
}