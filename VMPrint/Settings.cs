/****************************************************************************/
/*                                                                          */
/*  Copyright (C) 2019                                                      */
/*          （株）ＮＴＴデータ                                              */
/*           第三金融事業本部 戦略ビジネス本部　システム企画担当            */
/*                                                                          */
/*  収容物  ＣＯＮＴＩＭＩＸＥ    VirtualPrinterDriver                      */
/*                                                                          */
/****************************************************************************/
/*--------------------------------------------------------------------------*/
/*〈対象業務名〉                                                            */
/*〈対象業務ＩＤ〉                                                          */
/*〈モジュール名〉                  Settings                                */
/*〈モジュールＩＤ〉                                                        */
/*〈モジュール通番〉                                                        */
/*--------------------------------------------------------------------------*/
/* ＜適応ＯＳ＞                     Windows 10 XXX                          */
/* ＜開発環境＞                     Microsoft Visual Studio 2017            */
/*--------------------------------------------------------------------------*/
/* ＜開発システム名＞               ＣＯＮＴＩＭＩＸＥ                      */
/* ＜開発システム番号＞                                                     */
/*--------------------------------------------------------------------------*/
/* ＜開発担当名＞                   システム企画担当                        */
/* ＜電話番号＞                     050-5546-2418                           */
/*--------------------------------------------------------------------------*/
/* ＜設計者名＞                     範                                      */
/* ＜設計年月日＞                   2019年02月19日　　　　　                */
/* ＜設計修正者名＞                                                         */
/* ＜設計修正年月日及び修正ＩＤ＞                                           */
/*--------------------------------------------------------------------------*/
/* ＜ソース作成者名＞               範                                      */
/* ＜ソース作成年月日＞             2019年02月19日　　　　　                */
/* ＜ソース修正者名＞                                                       */
/* ＜ソース修正年月日及び修正ＩＤ＞                                         */
/*--------------------------------------------------------------------------*/
namespace VMPrint.Properties {
    
    
    // This class allows you to handle specific events on the settings class:
    //  The SettingChanging event is raised before a setting's value is changed.
    //  The PropertyChanged event is raised after a setting's value is changed.
    //  The SettingsLoaded event is raised after the setting values are loaded.
    //  The SettingsSaving event is raised before the setting values are saved.
    internal sealed partial class Settings {
        
        public Settings() {
            // // To add event handlers for saving and changing settings, uncomment the lines below:
            //
            // this.SettingChanging += this.SettingChangingEventHandler;
            //
            // this.SettingsSaving += this.SettingsSavingEventHandler;
            //
        }
        
        private void SettingChangingEventHandler(object sender, System.Configuration.SettingChangingEventArgs e) {
            // Add code to handle the SettingChangingEvent event here.
        }
        
        private void SettingsSavingEventHandler(object sender, System.ComponentModel.CancelEventArgs e) {
            // Add code to handle the SettingsSaving event here.
        }
    }
}
