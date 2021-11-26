using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoUpdate
{
	/// <summary>
	/// 创建一个实例
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public class SysT<T> where T : class, new()
	{
		public static T GetInstance
		{
			get
			{
				bool flag = SysT<T>.GetT == null;
				if (flag)
				{
					object @lock = SysT<T>.m_lock;
					lock (@lock)
					{
						bool flag3 = SysT<T>.GetT == null;
						if (flag3)
						{
							SysT<T>.GetT = Activator.CreateInstance<T>();
						}
					}
				}
				return SysT<T>.GetT;
			}
		}

		public static T GetT;

		public static readonly object m_lock = new object();
	}
}
