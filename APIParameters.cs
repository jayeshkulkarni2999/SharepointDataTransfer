using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

//using Xamarin.Forms;

namespace SharepointTransfer
{
    public class APIParameters
    {
        public string _PropertyName;
        public string _PropertyValue;

        public APIParameters()
        {
            _PropertyName = string.Empty;
            _PropertyValue = string.Empty;
        }

        public string PropertyName
        {
            get { return _PropertyName; }
            set { _PropertyName = value; }
        }

        public string PropertyValue
        {
            get { return _PropertyValue; }
            set { _PropertyValue = value; }
        }

    }

}