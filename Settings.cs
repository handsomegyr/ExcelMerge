using System.ComponentModel;

namespace ExcelApplication1
{
    public class Settings: INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, e);
        }

        private string activity_id;
        public string ActivityId
        {
            get { return activity_id; }
            set
            {
                activity_id = value;
                OnPropertyChanged(new PropertyChangedEventArgs("ActivityId"));
            }
        }

        private string prize_id;
        public string PrizeId
        {
            get { return prize_id; }
            set
            {
                prize_id = value;
                OnPropertyChanged(new PropertyChangedEventArgs("PrizeId"));
            }
        }


        private bool is_used;
        public bool IsUsed
        {
            get { return is_used; }
            set
            {
                is_used = value;
                OnPropertyChanged(new PropertyChangedEventArgs("IsUsed"));
            }
        }

        private string prjcode;
        public string Prjcode
        {
            get { return prjcode; }
            set
            {
                prjcode = value;
                OnPropertyChanged(new PropertyChangedEventArgs("Prjcode"));
            }
        }

        private string targetPath;
        public string TargetPath
        {
            get { return targetPath; }
            set
            {
                targetPath = value;
                OnPropertyChanged(new PropertyChangedEventArgs("PrizeId"));
            }
        }

    }
}
