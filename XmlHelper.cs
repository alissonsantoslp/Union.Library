using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Union.Library
{
    public class XmlHelper
    {
        public string ReadSettingAppConfig(string key)
        {
            return ConfigurationManager.AppSettings[key];
        }

        public string ReadConnectionString(string key)
        {
            ConnectionStringSettings conexao = ConfigurationManager.ConnectionStrings[key];
            return conexao.ConnectionString;
        }

        public void SaveAppSettings(string key, string valor)
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            KeyValueConfigurationCollection confsettings = config.AppSettings.Settings;
            confsettings[key].Value = valor;
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(config.AppSettings.SectionInformation.Name);
        }
    }
}
