/*
Navicat MySQL Data Transfer

Source Server         : localhost_3306
Source Server Version : 50728
Source Host           : localhost:3306
Source Database       : smart_monitor

Target Server Type    : MYSQL
Target Server Version : 50728
File Encoding         : 65001

Date: 2026-05-07 17:02:46
*/

SET FOREIGN_KEY_CHECKS=0;

-- ----------------------------
-- Table structure for model_config
-- ----------------------------
DROP TABLE IF EXISTS `model_config`;
CREATE TABLE `model_config` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `model_name` varchar(100) NOT NULL,
  `field1` varchar(255) DEFAULT '',
  `field2` varchar(255) DEFAULT '',
  `field3` varchar(255) DEFAULT '',
  `field4` varchar(255) DEFAULT '',
  `field5` varchar(255) DEFAULT '',
  `field6` varchar(255) DEFAULT '',
  `field7` varchar(255) DEFAULT '',
  `field8` varchar(255) DEFAULT '',
  `field9` varchar(255) DEFAULT '',
  `field10` varchar(255) DEFAULT '',
  `field11` varchar(255) DEFAULT '',
  `field12` varchar(255) DEFAULT '',
  `field13` varchar(255) DEFAULT '',
  `field14` varchar(255) DEFAULT '',
  `field15` varchar(255) DEFAULT '',
  `field16` varchar(255) DEFAULT '',
  `field17` varchar(255) DEFAULT '',
  `field18` varchar(255) DEFAULT '',
  `field19` varchar(255) DEFAULT '',
  `field20` varchar(255) DEFAULT '',
  `create_time` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=utf8mb4;

-- ----------------------------
-- Records of model_config
-- ----------------------------
INSERT INTO `model_config` VALUES ('5', 'MODLE222', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', '期限（年）', 'des', 'BZR', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:33:30');
INSERT INTO `model_config` VALUES ('6', 'MODLE123', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', 'FXDJ', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:33:46');
INSERT INTO `model_config` VALUES ('7', 'MODLE333', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', '期限（年）', 'des', 'BZR', 'ASFA', 'ASFGH', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:18');
INSERT INTO `model_config` VALUES ('8', 'KG678', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', 'des', 'BZR', 'OPOP1', 'ASFGH', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:43');
INSERT INTO `model_config` VALUES ('9', 'MODLE555', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', '期限（年）', 'des', 'BZR', 'ASFGH', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:44:10');
INSERT INTO `model_config` VALUES ('10', '科技模型6', 'JGMC', 'JGBM', 'DKBH', 'DKMC', 'DKYE', '期限（年）', 'des', 'BZR', 'ASFA', 'ASFGH', 'ASF', 'QWRQ', 'ASFR', '', '', '', '', '', '', '', '2026-05-07 16:44:34');

-- ----------------------------
-- Table structure for model_data
-- ----------------------------
DROP TABLE IF EXISTS `model_data`;
CREATE TABLE `model_data` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `model_name` varchar(100) NOT NULL,
  `field1` varchar(255) DEFAULT '',
  `field2` varchar(255) DEFAULT '',
  `field3` varchar(255) DEFAULT '',
  `field4` varchar(255) DEFAULT '',
  `field5` varchar(255) DEFAULT '',
  `field6` varchar(255) DEFAULT '',
  `field7` varchar(255) DEFAULT '',
  `field8` varchar(255) DEFAULT '',
  `field9` varchar(255) DEFAULT '',
  `field10` varchar(255) DEFAULT '',
  `field11` varchar(255) DEFAULT '',
  `field12` varchar(255) DEFAULT '',
  `field13` varchar(255) DEFAULT '',
  `field14` varchar(255) DEFAULT '',
  `field15` varchar(255) DEFAULT '',
  `field16` varchar(255) DEFAULT '',
  `field17` varchar(255) DEFAULT '',
  `field18` varchar(255) DEFAULT '',
  `field19` varchar(255) DEFAULT '',
  `field20` varchar(255) DEFAULT '',
  `create_time` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=70 DEFAULT CHARSET=utf8mb4;

-- ----------------------------
-- Records of model_data
-- ----------------------------
INSERT INTO `model_data` VALUES ('1', 'MODLE123', '某银行123', 'TXM123457HG', 'DK223', '超额贷款', '2234.23', '2', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:03:01');
INSERT INTO `model_data` VALUES ('2', 'MODLE123', '某银行123', 'TXM123457HG', 'DK223', '超额贷款', '12.05', '3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:03:01');
INSERT INTO `model_data` VALUES ('3', 'MODLE123', '大地银行', 'TXM0981203J', 'DK112', '普通贷款', '2000.0', '1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:03:01');
INSERT INTO `model_data` VALUES ('43', 'MODLE222', 'KFX银行', 'LFXK123YU', 'YT001', '闪电贷', '8000000', '12.0', 'xx1', '刘大风', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:30:08');
INSERT INTO `model_data` VALUES ('44', 'MODLE222', '投行123', 'TX0000001', 'KLJ8YU', '小额贷款', '5000', '0.5', '阿斯蒂芬', 'Kcreis', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:30:08');
INSERT INTO `model_data` VALUES ('45', 'MODLE222', '投行123', 'TX0000001', 'KLJ8YU', '小额贷款', '200', '1.0', 'nan', '大华', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:30:08');
INSERT INTO `model_data` VALUES ('57', 'MODLE222', '投行123', 'TX0000001', 'KLJ8YU', '小额贷款', '200', '1', 'nan', '大华', '', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 11:59:40');
INSERT INTO `model_data` VALUES ('58', 'MODLE333', 'A123', 'AVC123', 'YT0W01', '贷款1', '22', '12', '打发', '风下', '12', '1', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:18');
INSERT INTO `model_data` VALUES ('59', 'MODLE333', 'A123', 'AVC123', 'YT0W01', '贷款1', '5000', '0.5', 'xxx123', '明飞', '2', 'nan', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:18');
INSERT INTO `model_data` VALUES ('60', 'MODLE333', 'A123', 'AVC123', 'YT0W01', '贷款1', '200', '1', 'xxxx', '二哥', 'nan', '2', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:18');
INSERT INTO `model_data` VALUES ('61', 'KG678', 'JG123', 'YU123', 'KG111', '欢乐贷1', '1', '123123132', '刘备', '嗷嗷', '1', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:43');
INSERT INTO `model_data` VALUES ('62', 'KG678', 'JG123', 'YU123', 'KG111', '欢乐贷1', '2', '1231231231231阿斯蒂芬', '马云', 'a', '2', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:43');
INSERT INTO `model_data` VALUES ('63', 'KG678', 'JG128', 'YU124', 'KG112', '欢乐贷2', '3', '的分公司电饭锅', 'desa', 'nan', '2', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:43:43');
INSERT INTO `model_data` VALUES ('64', 'MODLE555', 'BBB23', 'BVB23', 'BBAA1', '卡牌2', '234', '1', 'aaaa', '风清扬', '1', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:44:10');
INSERT INTO `model_data` VALUES ('65', 'MODLE555', 'BBB23', 'BVB23', 'BBAA1', '卡牌2', '1252', '2', 'asdfb', '宇乘', 'nan', '', '', '', '', '', '', '', '', '', '', '', '2026-05-07 16:44:10');
INSERT INTO `model_data` VALUES ('66', '科技模型6', '招牌银行123', 'ZP123', 'ZP22467', 'ZP贷款123', '22', '12', '打发', '风下', '12', '1', '1', 'nan', '78', '', '', '', '', '', '', '', '2026-05-07 16:44:34');
INSERT INTO `model_data` VALUES ('67', '科技模型6', '招牌银行123', 'ZP123', 'ZP22467', 'ZP贷款123', '5000', '2。5', 'basdf', '明飞', '2', 'nan', 'nan', '2', '8', '', '', '', '', '', '', '', '2026-05-07 16:44:34');
INSERT INTO `model_data` VALUES ('68', '科技模型6', '招牌银行123', 'ZP123', 'ZP22469', 'ZP贷款222', '100', '1', '阿斯顿发', '市长', 'nan', '2', 'nan', 'nan', '9', '', '', '', '', '', '', '', '2026-05-07 16:44:34');
INSERT INTO `model_data` VALUES ('69', '科技模型6', '招牌银行123', 'ZP123', 'ZP22468', 'ZP贷款235', '12', '2', '保罗', '保罗', 'nan', '2', 'nan', 'nan', 'nan', '', '', '', '', '', '', '', '2026-05-07 16:44:34');

-- ----------------------------
-- Table structure for sys_config
-- ----------------------------
DROP TABLE IF EXISTS `sys_config`;
CREATE TABLE `sys_config` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `db_host` varchar(50) DEFAULT NULL,
  `db_user` varchar(50) DEFAULT NULL,
  `db_pwd` varchar(255) DEFAULT NULL,
  `db_name` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- ----------------------------
-- Records of sys_config
-- ----------------------------
