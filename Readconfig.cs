using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoUpdate
{
	internal class Readconfig : SysT<Readconfig>
	{
		public Readconfig()
		{
			this.dictionary = new Dictionary<string, string>();
		}
		public void ReadConfig(string path)
		{
			bool flag = !File.Exists(path);
			if (!flag)
			{
				string[] array = File.ReadAllLines(path, System.Text.Encoding.UTF8);
				for (int i = 0; i < array.Length; i++)
				{
					bool flag2 = array[i].Split(new char[]
					{
						'='
					}).Length == 2;
					if (flag2)
					{
						string key = array[i].Split(new char[]
						{
							'='
						})[0];
						string value = array[i].Split(new char[]
						{
							'='
						})[1];
						bool flag3 = !this.dictionary.ContainsKey(key);
						if (flag3)
						{
							this.dictionary.Add(key, value);
						}
					}
				}
				this.IsReaded = true;
				this.Count = dictionary.Count;
			}
		}

		public string this[string key]
		{
			get
			{
				bool flag = this.dictionary.ContainsKey(key);
				string result;
				if (flag)
				{
					result = this.dictionary[key];
				}
				else
				{
					result = null;
				}
				return result;
			}
			set
			{
				bool flag = this.dictionary.ContainsKey(key);
				if (flag)
				{
					this.dictionary[key] = value;
				}
				else
				{
					this.dictionary.Add(key, value);
				}
			}
		}
		private Dictionary<string, string> dictionary;

		public bool IsReaded;
		public int Count;
	}
}
