using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Drawing;
using System.Windows.Forms;


namespace MMIT.ShareCleaner
{


    public class Node : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string FullPathFi;//Full Path of a File
        public string FullPathDi;//Full Path of a Folder
        public string PathFi;//name of a File
        public string PathDi;//name of a Folder
        public event PropertyChangedEventHandler PropertyChanged;

        private bool isDeleteChecked;
        private bool isArchiveChecked;
        private bool isIgnoreChecked =true;
        private float size;
        private DateTime lastAccess;
        private string background;
        private bool isEnabled = true;



        //Databinding
        public bool IsEnabled
        {
            
            get { return isEnabled; }
            set
            {
                if (isEnabled != value)
                {
                    isEnabled = value;
                    OnPropertyChanged("IsEnabled");
                }
            }
        }
        public string Background
        {
            get { return background; }
            set
            {
                if (background != value)
                {
                    background = value;
                    OnPropertyChanged("Background");
                }
            }
        }

        public DateTime LastAccess
        {
            get { return lastAccess; }
            set
            {
                if (lastAccess != value)
                {
                    lastAccess = value;
                    OnPropertyChanged("LastAccess");
                }
            }
        }
        public float Size
        {
            get { return size; }
            set
            {
                if (size != value)
                {
                    size = value;
                    OnPropertyChanged("Size");
                }
            }
        }
        public bool IsDeleteChecked
        {
            get { return isDeleteChecked; }
            set
            {
                if (isDeleteChecked != value)
                {
                    isDeleteChecked = value;
                    OnPropertyChanged("IsDeleteChecked");
                }
            }
        }
        public bool IsArchiveChecked
        {
            get { return isArchiveChecked; }
            set
            {
                if (isArchiveChecked != value)
                {
                    isArchiveChecked = value;
                    OnPropertyChanged("IsArchiveChecked");
                }
            }
        }
        public bool IsIgnoreChecked
        {
            get { return isIgnoreChecked; }
            set
            {
                if (isIgnoreChecked != value)
                {
                    isIgnoreChecked = value;
                    OnPropertyChanged("IsIgnoreChecked");
                }
            }
        }
        //Updating the UI
        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        //Generating the Treeview
        public List<Node> Children { get; } = new List<Node>();

        public static Node GetTree(DirectoryInfo di)
        {
          
                Node item = new Node();
                item.lastAccess = di.LastAccessTime;
                item.PathDi = di.Name;
           

                item.FullPathDi = di.FullName;
                item.Name = di.Name;
            
                foreach (DirectoryInfo s in di.GetDirectories())
                {
                    item.Children.Add(GetTree(s));
                }
                foreach (FileInfo fi in di.GetFiles())
                {
                    item.Children.Add(new Node { Name = fi.Name, size = fi.Length, lastAccess = fi.LastAccessTime, FullPathFi = fi.FullName, PathFi = fi.Name });
                }
                return item;
           
            }
        //Setting the Checked status of all children



        public void CheckAllChildrenDelete(bool IsDeleteChecked)
        {
            foreach (Node item in Children)
            {
                item.IsDeleteChecked = IsDeleteChecked;
                if (IsDeleteChecked == true)
                {
                    item.IsEnabled = false;
                }
                if (IsDeleteChecked == false)
                {
                    item.IsEnabled = true;
                }

                item.CheckAllChildrenDelete(IsDeleteChecked);
            }
            
        }
        public void CheckAllChildrenArchive(bool IsArchiveChecked)
        {
            foreach (Node item in Children)
            {
                item.IsArchiveChecked = IsArchiveChecked;
                if (IsArchiveChecked == true)
                {
                    item.IsEnabled = false;
                }
                if (IsArchiveChecked == false)
                {
                    item.IsEnabled = true;
                }

                item.CheckAllChildrenArchive(IsArchiveChecked);
            }
        }
        public void CheckAllChildrenIgnore(bool IsIgnoreChecked)
        {
            foreach (Node item in Children)
            {
                item.IsIgnoreChecked = IsIgnoreChecked;
                item.IsEnabled = true;

                item.CheckAllChildrenIgnore(IsIgnoreChecked);
            }
        }
    }
}



