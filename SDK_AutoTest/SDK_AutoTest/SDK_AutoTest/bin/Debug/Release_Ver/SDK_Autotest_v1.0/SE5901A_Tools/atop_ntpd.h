typedef enum timezone{
	GMT_MINUS12_00=0,				//-- (GMT-12:00) International Date Line West
	GMT_MINUS11_00,					//-- (GMT-11:00) Midway Island, Samoa
	GMT_MINUS10_00,					//-- (GMT-10:00) Hawaii
	GMT_MINUS09_00,					//-- (GMT-09:00) Alaska
	GMT_MINUS08_00,					//-- (GMT-08:00) Pacific Time (US & Cannada)
	GMT_MINUS07_00,					//-- (GMT-07:00) Arizona, Chihuahua, La Paz, Mazatlan, Mountain Time (US & Canada)
	GMT_MINUS06_00=8,				//-- (GMT-06:00) Central America, Central Time (US & Canada), Guadalajara, Mezico City, Monterrey, Saskatchewan
	GMT_MINUS05_00=12,			//-- (GMT-05:00) Bogota,Lima,Quito,Eastern Time (US & Canada), Indiana (East)
	GMT_MINUS04_00=15,			//-- (GMT-04:00) Atlantic Time (Canada), Caracas, La Paz, Santiago
	GMT_MINUS03_30=18,			//-- (GMT-03:30) Newfoundland
	GMT_MINUS03_00,					//-- (GMT-03:00) Brasilia, Buenos Aires, Georgetown
	GMT_MINUS02_00=21,			//-- (GMT-02:00) Mid-Atlantic
	GMT_MINUS01_00,					//-- (GMT-01:00) Azores, Cape Verde Is.
	GMT_PLUS00_00,					//-- (GMT) Casablanca, Monrovia, Greenwich Mean Time: Dublin, Edinburgh, Lisbon, London
	GMT_PLUS01_00=25,				//-- (GMT+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna, Belgrade, Bratislava, Budapest, Ljubljana, Prague, Brussels, Copenhagen, Madrid, Paris, Vilnius, Sarajevo, Skopje, Sofija, Warsaw, Zagreb
	GMT_PLUS02_00=29,				//-- (GMT+02:00) Athens, Istanbul, Minsk, Bucharest, Cairo, Harare, Pretoria, Helsinki, Riga, Tallinn, Jerusalem
	GMT_PLUS03_00=35,				//-- (GMT+03:00) Baghdad, Kuwait, Riyadh, Moscow, St. Petersburg, Volgograd, Nairobi
	GMT_PLUS03_30=38,				//-- (GMT+03:30) Tehran 
	GMT_PLUS04_00,					//-- (GMT+04:00) Abu Dhabi, Muscat,Baku, Tbilisi, Yerevan
	GMT_PLUS04_30=41,				//-- (GMT+04:30) Kabul
	GMT_PLUS05_00,					//-- (GMT+05:00) Ekaterinburg, Islamabad, Karachi, Tashkent
	GMT_PLUS05_30=44,				//-- (GMT+05:30) Bombay, Calcutta, Madras, New Delhi
	GMT_PLUS06_00,					//-- (GMT+06:00) Astana, Dhaka, Colombo
	GMT_PLUS07_00=47,				//-- (GMT+07:00) Bangkok, Hanoi, Jakarta
	GMT_PLUS08_00,					//-- (GMT+08:00) Beijing, Chongqing, Hong Kong, Urumqi, Perth, Singapore, Taipei
	GMT_PLUS09_00=52,				//-- (GMT+09:00) Osaka, Sapporo, Tokyo, Seoul, Yakutsk
	GMT_PLUS09_30=55,				//-- (GMT+09:30) Adelaide, Darwin
	GMT_PLUS10_00=57,				//-- (GMT+10:00) Brisbane, Canberra, Melbourne, Sydney, Guam, Port Moresby, Hobart, Vladivostok"
	GMT_PLUS11_00=62,				//-- (GMT+11:00) Auckland, Wellington, Magadan, Solamon, New Caledonia
	GMT_PLUS12_00=64				//-- (GMT+12:00) Fiji, Kamchatka, Marshall Is.
}TIMEZONE;

typedef enum saving_time_momth{
	MONTH_JAN=1,				//-- Jan
	MONTH_FEB,					//-- Feb			
	MONTH_MAR,					//-- Mar	
	MONTH_APR,					//-- Apr	
	MONTH_MAY,					//-- May	
	MONTH_JUN,					//-- Jun	
	MONTH_JUL,					//-- Jul	
	MONTH_AUG,					//-- Aug	
	MONTH_SEP,					//-- Sep	
	MONTH_OCT,					//-- Oct	
	MONTH_NOV,					//-- Nov	
	MONTH_DEC						//-- Dec
}SAVING_TIME_MONTH;

typedef enum saving_time_week{
	WEEK_1ST=1,				//-- 1st
	WEEK_2ND,					//-- 2nd
	WEEK_3RD,					//-- 3rd
	WEEK_4TH,					//-- 4th
	WEEK_LAST					//-- last
}SAVING_TIME_WEEK;

typedef enum saving_time_day{
	DAY_SUN=0,				//-- Sun
	DAY_MON,					//-- Mon
	DAY_TUE,					//-- Tue
	DAY_WED,					//-- Wed
	DAY_THU,					//-- Thu
	DAY_FRI,					//-- Fri
	DAY_SAT 					//-- Sat
}SAVING_TIME_DAY;

typedef enum saving_time_hour{
	HOUR_00=0,				
	HOUR_01,
	HOUR_02,
	HOUR_03,
	HOUR_04,					
	HOUR_05,
	HOUR_06,
	HOUR_07,
	HOUR_08,
	HOUR_09,
	HOUR_10,
	HOUR_11,
	HOUR_12,
	HOUR_13,
	HOUR_14,
	HOUR_15,
	HOUR_16,
	HOUR_17,
	HOUR_18,
	HOUR_19,					
	HOUR_20,					
	HOUR_21,					
	HOUR_22,					
	HOUR_23 					
}SAVING_TIME_HOUR;

typedef enum saving_time_offset{
	OFFSET_00=0,		
	OFFSET_01,
	OFFSET_02,
	OFFSET_03,
	OFFSET_04,					
	OFFSET_05,
	OFFSET_06,
	OFFSET_07,
	OFFSET_08,
	OFFSET_09,
	OFFSET_10,
	OFFSET_11,
	OFFSET_12					
}SAVING_TIME_OFFSET;