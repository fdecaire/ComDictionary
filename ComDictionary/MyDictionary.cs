using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ComDictionary
{
	[InterfaceType(ComInterfaceType.InterfaceIsDual),
	Guid("8075919D-4153-4277-8BB1-5F1BB6EF6588")] 
	public interface IMyDictionary
	{
		object this[string key] { get; set; }
		void Clear();
	}

	[ClassInterface(ClassInterfaceType.None),
	Guid("DF9B2D9F-18A5-42A1-8B64-520ADF861BAF"), ProgId("ComDictionary.MyDictionary")]
    public class MyDictionary : IMyDictionary
    {
		private Dictionary<string, object> Vars = new Dictionary<string, object>();

		public object this[string key]
		{
			get 
			{
				if (Vars.ContainsKey(key.ToLower()))
				{
					return Vars[key.ToLower()];
				}
				else
				{
					return null;
				}
			}
			set 
			{
				Vars[key.ToLower()] = value;
			}
		}

		public void Clear()
		{
			Vars.Clear();
		}
    }
}
