CREATE TABLE ds_userTBL (
	ds_user_id			CHAR(25)	PRIMARY KEY,
	ds_user_password	CHAR(25)	NOT NULL,
	ds_user_name		CHAR(10)	NOT NULL,
	ds_user_is_save		BIT			NOT NULL DEFAULT 0	,
	ds_user_is_auto		BIT			NOT NULL DEFAULT 0		
)

SELECT * 
	--DELETE
	FROM ds_userTBL