# 前置说明
- 提取数据的checklist.xlsx为客户释放，工作表的内容区域做了保护，我们无法更多编辑，只能正常使用；
- 针对不同的network、codec、bandwidth，checklist的内容、测试的具体项目会有所区别，为了方便工作，继承到了一份checklist中，然后根据不同的场景，动态切换到对应的工作表（手动切换）。
- HA（handset）手持、HE（Headset）耳机、HH（handsfree）免提，有不同的checklist，从当前版本的模板来说，分别对应`Voice_Tuning_Checklist_v5.0.2_Handset.xlsx`、`Voice_Tuning_Checklist_v5.0.2_Headset.xlsx`、`Voice_Tuning_Checklist_v5.0.2_Handsfree.xlsx`;

# 动态切换工作表说明

## HA(Handset)
在`Voice_Tuning_Checklist_v5.0.2_Handset.xlsx`中，在`Report`这个sheet页里面，有一个基本信息面板；
- B13 这个单元格是`Headser Interface`的内容，这个单元格支持下来选择，选项内容包含：
  - 3.5mm Analog + USB-C Digital
  - USB-C Analog + Digital
  - USB-C Digital only
15行是Test Network的选项，其中：
- B15 是`Network`，这个单元格支持下拉选择，选项内容包括：
  - GSM
  - WCDMA
  - CDMA
  - VoLTE
  - VolP
  - VoNR
  - VoWiFi
- C15 是`Vocoder`，这个单元格支持下拉选择，选项内容包括:
  - AMR_NB
  - AMR_WB
  - EVS_NB
  - EVS_WB
  - EVS_SWB
- D15 是`Bitrate`的选项，这个单元格支持下拉选择，选项内容包括：
  - 9.6k/bits
  - 13.2k/bits
  - 128k/bits

然后在第二个sheet页也就是`Handset`这个sheet页，就是详细的测试项目的checklist；


