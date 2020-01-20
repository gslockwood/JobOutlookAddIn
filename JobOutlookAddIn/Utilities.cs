using System;
using System.Collections.Generic;

namespace Utilities
{
	public enum SendCondition
	{
		Conditional,
		Force,
		Stop
	}

	/// <summary>
	/// Summary description for FileUtilities.
	/// </summary>
	public class FileUtilities
	{
		public static string ReadFile(string fileName )
		{
			return System.IO.File.ReadAllText(fileName);
		}

		internal static IList<Entity> ParseFile(string fileName)
		{
			IList<Entity> entityList = new List<Entity>();
			try
			{
				foreach (string line in System.IO.File.ReadAllLines(fileName))
				{
					if( (string.IsNullOrEmpty(line)) || (line[0] == ';') )
						continue;

					String[] array = line.Split( ',' );
					if (array.Length < 2)
						continue;

					entityList.Add(new Entity(array[0], array[1]));
					//
				}
				//
			}
			catch (Exception e)
			{
			}

			return entityList;
			//
		}
		//
	}
	//
}

