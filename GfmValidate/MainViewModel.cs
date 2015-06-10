
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace GfmValidate
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public string GitHubUsername
        {
            get
            {
                return Properties.Settings.Default.GitHubUsername;
            }

            set
            {
                if (value != Properties.Settings.Default.GitHubUsername)
                {
                    Properties.Settings.Default.GitHubUsername = value;
                    RaisePropertyChanged();
                }
            }
        }

        public string GitHubPassword
        {
            get
            {
                return Properties.Settings.Default.GitHubPassword;
            }

            set
            {
                if (value != Properties.Settings.Default.GitHubPassword)
                {
                    Properties.Settings.Default.GitHubPassword = value;
                    RaisePropertyChanged();
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void RaisePropertyChanged([CallerMemberName] string propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
