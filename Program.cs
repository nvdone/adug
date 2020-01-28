//NVD ActiveDirectoryUsersGroups
//Copyright © 2016, Nikolay Dudkin

//This program is free software: you can redistribute it and/or modify
//it under the terms of the GNU General Public License as published by
//the Free Software Foundation, either version 3 of the License, or
//(at your option) any later version.
//This program is distributed in the hope that it will be useful,
//but WITHOUT ANY WARRANTY; without even the implied warranty of
//MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
//GNU General Public License for more details.
//You should have received a copy of the GNU General Public License
//along with this program.If not, see<https://www.gnu.org/licenses/>.

using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace adug
{
	class Program
	{
		static void Main(string[] args)
		{
			Console.Out.WriteLine(string.Format("{0} {1}\r\n{2}", ((AssemblyTitleAttribute)Attribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly(), typeof(AssemblyTitleAttribute))).Title, Assembly.GetExecutingAssembly().GetName().Version, ((AssemblyCopyrightAttribute)Attribute.GetCustomAttribute(System.Reflection.Assembly.GetExecutingAssembly(), typeof(AssemblyCopyrightAttribute))).Copyright));

			if (args.Length < 2)
			{
				Console.Out.WriteLine("\r\nUsage: adug.exe domain_name output_file.xlsx [\"parent OU distinguished name\"]");
				return;
			}

			Console.Out.WriteLine();
			ToHere();
			To("Loading...");

			List<string> groupNames = new List<string>();

			Dictionary<string, List<string>> userNamesGroupNames = new Dictionary<string, List<string>>();

			List<string> userNames = getAllActiveDomainUserNames(args.Length == 3 ? args[2] : args[0]);
			userNames.Sort();

			int userNamesDone = 0;

			To("Loading: ");
			ToHere();
			To("0%");

			foreach (string userName in userNames)
			{
				List<string> userGroupNames = getUserGroupNames(args[0], userName);
				userGroupNames.Sort();

				userNamesGroupNames.Add(userName, userGroupNames);

				foreach (string userGroupName in userGroupNames)
				{
					if (!groupNames.Contains(userGroupName, StringComparer.InvariantCultureIgnoreCase))
						groupNames.Add(userGroupName);
				}

				To(string.Format("{0}%", (int)(((double)++userNamesDone) / userNames.Count * 100)));
			}

			groupNames.Sort();

			To("done.");

			if (userNames.Count > 1000000)
			{
				Console.Out.WriteLine("Error: too many users!");
				return;
			}

			if (groupNames.Count > 20000)
			{
				Console.Out.WriteLine("Error: too many groups!");
				return;
			}

			Console.Out.WriteLine();
			ToHere();
			To("Exporting...");

			ExcelPackage excelPackage = new ExcelPackage();
			ExcelWorksheet ws = excelPackage.Workbook.Worksheets.Add("ADUG");

			ws.Cells[1, 1].Value = "Login";
			ws.Cells[1, 2].Value = "Display name";
			for (int i = 0; i < groupNames.Count; i++)
			{
				ws.Cells[1, i + 3].Value = args[0].Contains(".") ? groupNames[i] + "@" + args[0] : args[0] + "\\" + groupNames[i];
				ws.Cells[1, i + 3].Style.WrapText = true;
			}

			for (int i = 0; i < groupNames.Count + 2; i++)
			{
				ws.Cells[1, i + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
				ws.Cells[1, i + 1].Style.Font.Bold = true;
				ws.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
				ws.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightSteelBlue);
				ws.Column(i + 1).Width = i < 2 ? 35 : 20;
			}

			ws.Cells[1, 1, 1, groupNames.Count + 2].AutoFilter = true;
			ws.View.FreezePanes(2, 3);

			int rowindex = 2;

			To("Exporting: ");
			ToHere();
			To("0%");

			foreach (string userName in userNamesGroupNames.Keys)
			{
				ws.Cells[rowindex, 1].Value = args[0].Contains(".") ? userName + "@" + args[0] : args[0] + "\\" + userName;
				ws.Cells[rowindex, 2].Value = getUserDisplayName(args[0], userName);

				for (int i = 0; i < groupNames.Count; i++)
				{
					if (userNamesGroupNames[userName].Contains(groupNames[i], StringComparer.InvariantCultureIgnoreCase))
					{
						ws.Cells[rowindex, i + 3].Value = "Y";
					}
				}

				for (int i = 0; i < groupNames.Count + 2; i++)
				{
					ws.Cells[rowindex, i + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
				}

				To(string.Format("{0}%", (int)(((double)++rowindex - 2) / userNames.Count * 100)));
			}

			using (FileStream fs = new FileStream(args[1], FileMode.Create, FileAccess.ReadWrite))
			{
				excelPackage.SaveAs(fs);
			}

			To("done.");
		}

		static int console_len = 0;
		static int console_row = 0;
		static int console_col = 0;

		static void ToHere()
		{
			console_len = 0;
			console_row = Console.CursorTop;
			console_col = Console.CursorLeft;
		}

		static void To(string str)
		{
			Console.SetCursorPosition(console_col, console_row);
			Console.Out.Write(str);
			Console.Write(new String(' ', Math.Max(0, console_len - str.Length)));
			console_len = str.Length;
		}

		static List<string> getAllActiveDomainUserNames(string domainName)
		{
			List<string> userNames = new List<string>();
			try
			{
				DirectoryEntry root = new DirectoryEntry("LDAP://" + domainName, null, null, AuthenticationTypes.Secure);

				if (root == null)
					return userNames;

				SearchResultCollection results;
				DirectorySearcher searcher = new DirectorySearcher(root);
				searcher.Filter = "(objectClass=User)";
				results = searcher.FindAll();

				foreach (SearchResult result in results)
				{
					if (!result.Properties["objectclass"].Contains("computer"))
					{
						if ((((int)result.Properties["useraccountcontrol"][0]) & 0x2) == 0)
							userNames.Add(result.Properties["samaccountname"][0].ToString());
					}
				}

				searcher.Dispose();
				root.Close();
				root.Dispose();
			}
			catch { }

			return userNames;
		}

		public static List<string> getUserGroupNames(string domain, string login)
		{
			List<string> groupNames = new List<string>();
			try
			{
				string filter = string.Format("(&(ObjectClass={0})(sAMAccountName={1}))", "person", login);

				DirectoryEntry root = new DirectoryEntry("LDAP://" + domain, null, null, AuthenticationTypes.Secure);

				if (root == null)
				{
					return groupNames;
				}

				DirectorySearcher searcher = new DirectorySearcher(root);
				if (searcher == null)
				{
					root.Close();
					root.Dispose();
					return groupNames;
				}

				searcher.SearchScope = SearchScope.Subtree;
				searcher.ReferralChasing = ReferralChasingOption.All;
				searcher.Filter = filter;
				SearchResult result = searcher.FindOne();
				if (result == null)
				{
					searcher.Dispose();
					root.Close();
					root.Dispose();
					return groupNames;
				}

				DirectoryEntry entry = result.GetDirectoryEntry();

				if (entry == null)
				{
					searcher.Dispose();
					root.Close();
					root.Dispose();
					return groupNames;
				}

				if (entry.Properties == null || entry.Properties.Count < 1)
				{
					entry.Close();
					entry.Dispose();
					searcher.Dispose();
					root.Close();
					root.Dispose();
					return groupNames;
				}

				string primaryGroupName = getPrimaryGroupProperty(entry, root, "samaccountname");
				if (primaryGroupName.Length > 0)
					groupNames.Add(primaryGroupName);

				attributeValuesMultiString("memberOf", getPrimaryGroupPath(entry, root), groupNames, true, "", new Regex(@"CN=([^,]*),"));

				attributeValuesMultiString("memberOf", entry.Path, groupNames, true, "", new Regex(@"CN=([^,]*),"));

				entry.Close();
				entry.Dispose();
				searcher.Dispose();
				root.Close();
				root.Dispose();
			}
			catch { }

			return groupNames;
		}

		static string getUserDisplayName(string domain, string login)
		{
			string displayName = "";

			try
			{
				string filter = string.Format("(&(ObjectClass={0})(sAMAccountName={1}))", "person", login);
				string[] properties = { "displayName" };

				DirectoryEntry root = new DirectoryEntry("LDAP://" + domain, null, null, AuthenticationTypes.Secure);

				if (root == null)
					return "";

				DirectorySearcher searcher = new DirectorySearcher(root);

				if (searcher == null)
				{
					root.Close();
					root.Dispose();
					return "";
				}

				searcher.SearchScope = SearchScope.Subtree;
				searcher.ReferralChasing = ReferralChasingOption.All;
				searcher.PropertiesToLoad.AddRange(properties);
				searcher.Filter = filter;

				if (searcher.FindOne() != null)
				{
					if (searcher.FindOne().Properties["displayName"].Count > 0)
						displayName = searcher.FindOne().Properties["displayName"][0].ToString();
				}

				searcher.Dispose();
				root.Close();
				root.Dispose();
			}
			catch { }

			return displayName;
		}

		static List<string> attributeValuesMultiString(string attributeName, string objectDN, List<string> valuesCollection, bool recursive, string prefix, Regex regex)
		{
			DirectoryEntry entry = new DirectoryEntry(objectDN);
			PropertyValueCollection valueCollection = entry.Properties[attributeName];
			IEnumerator enumerator = valueCollection.GetEnumerator();

			while (enumerator.MoveNext())
			{
				if (enumerator.Current != null)
				{
					string groupName = prefix + enumerator.Current.ToString();

					if (regex != null)
					{
						if (regex.IsMatch(enumerator.Current.ToString()))
						{
							groupName = prefix + regex.Match(enumerator.Current.ToString()).Groups[1].Value;
						}
						else
							continue;
					}

					if (!valuesCollection.Contains(groupName))
					{
						valuesCollection.Add(groupName);

						if (recursive)
						{
							attributeValuesMultiString(attributeName, "LDAP://" +
							enumerator.Current, valuesCollection, true, prefix, regex);
						}
					}
				}
			}

			entry.Close();
			entry.Dispose();
			return valuesCollection;
		}

		static string getPrimaryGroupProperty(DirectoryEntry entry, DirectoryEntry domainEntry, string propertyName)
		{
			string retval = "";

			try
			{
				int primaryGroupId = (int)entry.Properties["primaryGroupID"].Value;
				byte[] objectSid = (byte[])entry.Properties["objectSid"].Value;

				System.Text.StringBuilder escapedGroupSid = new System.Text.StringBuilder();

				for (int i = 0; i < objectSid.Length - 4; i++)
				{
					escapedGroupSid.AppendFormat("\\{0:x2}", objectSid[i]);
				}

				for (int i = 0; i < 4; i++)
				{
					escapedGroupSid.AppendFormat("\\{0:x2}", (primaryGroupId & 0xFF));
					primaryGroupId >>= 8;
				}

				DirectorySearcher searcher = new DirectorySearcher();
				if (domainEntry != null)
				{
					searcher.SearchRoot = domainEntry;
				}

				searcher.Filter = "(&(objectCategory=Group)(objectSID=" + escapedGroupSid.ToString() + "))";
				searcher.PropertiesToLoad.Add(propertyName);

				if (searcher.FindOne() != null)
				{
					if (searcher.FindOne().Properties[propertyName].Count > 0)
						retval = searcher.FindOne().Properties[propertyName][0].ToString();
				}

				searcher.Dispose();
			}
			catch { }

			return retval;
		}

		static string getPrimaryGroupPath(DirectoryEntry entry, DirectoryEntry domainEntry)
		{
			string retval = "";

			try
			{
				int primaryGroupId = (int)entry.Properties["primaryGroupID"].Value;
				byte[] objectSid = (byte[])entry.Properties["objectSid"].Value;

				System.Text.StringBuilder escapedGroupSid = new System.Text.StringBuilder();

				for (uint i = 0; i < objectSid.Length - 4; i++)
				{
					escapedGroupSid.AppendFormat("\\{0:x2}", objectSid[i]);
				}

				for (uint i = 0; i < 4; i++)
				{
					escapedGroupSid.AppendFormat("\\{0:x2}", (primaryGroupId & 0xFF));
					primaryGroupId >>= 8;
				}

				DirectorySearcher searcher = new DirectorySearcher();
				if (domainEntry != null)
				{
					searcher.SearchRoot = domainEntry;
				}

				searcher.Filter = "(&(objectCategory=Group)(objectSID=" + escapedGroupSid.ToString() + "))";

				SearchResult result = searcher.FindOne();
				if (result != null)
				{
					DirectoryEntry directoryEntry = result.GetDirectoryEntry();

					retval = directoryEntry.Path;

					directoryEntry.Close();
					directoryEntry.Dispose();
				}

				searcher.Dispose();
			}
			catch { }

			return retval;
		}
	}
}
