CREATE TABLE item_m (
  item_id bigint(20) unsigned NOT NULL COMMENT 'アイテムID',
  item_name varchar(64) NOT NULL COMMENT 'アイテム名',
  effect_type int(11) unsigned NOT NULL COMMENT '効果タイプ',
  effect_value1 varchar(16) NULL COMMENT '効果値1',
  effect_value2 varchar(16) NULL COMMENT '効果値2',
  flavor_text varchar(256) NULL COMMENT 'フレーバーテキスト',
  image_resource varchar(128) NULL COMMENT 'リソース用文字列',
  PRIMARY KEY (`item_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8 COMMENT='アイテムマスタ';
