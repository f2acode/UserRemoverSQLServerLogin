// -----------------------------------------------------------------------
// <copyright file="MainWindow.xaml.cs" company="Mark Walker">
//     Copyright © Mark Walker 2013. All rights reserved.
// </copyright>
// -----------------------------------------------------------------------
namespace SSMSMRU
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.Serialization.Formatters.Binary;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;

    using Microsoft.SqlServer.Management.UserSettings;

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        #region Fields

        /// <summary>List of servers</summary>
        private Dictionary<string, Server> servers;

        /// <summary>Settings file path</summary>
        private string settingsFilePath;

        /// <summary>Sql studio settings</summary>
        private SqlStudio sqlStudioSettings;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        /// Gets or sets a value indicating whether this instance has changes ready to be saved.
        /// </summary>
        internal bool HasChanges
        {
            get
            {
                return this.uxSave.IsEnabled;
            }

            set
            {
                this.uxSave.IsEnabled = value;
                this.uxCancel.IsEnabled = value;
            }
        }

        /// <summary>
        /// Gets the servers.
        /// </summary>
        private Dictionary<string, Server> Servers
        {
            get
            {
                if (this.servers == null)
                {
                    this.servers = new Dictionary<string, Server>();
                    var binaryFormatter = new BinaryFormatter();
                    var inStream = new MemoryStream(File.ReadAllBytes(this.SettingsFilePath));
                    this.sqlStudioSettings = (SqlStudio)binaryFormatter.Deserialize(inStream);
                    foreach (var pair in this.sqlStudioSettings.SSMS.ConnectionOptions.ServerTypes)
                    {
                        ServerTypeItem serverTypeItem = pair.Value;

                        foreach (ServerConnectionItem serverConnectionItem in serverTypeItem.Servers)
                        {
                            var serverName = GetServerName(serverConnectionItem, this.servers);
                            this.servers.Add(serverName, new Server(serverConnectionItem, serverName));
                        }
                    }
                }

                return this.servers;
            }
        }

        /// <summary>
        /// Gets the settings file path.
        /// </summary>
        private string SettingsFilePath
        {
            get
            {
                if (!File.Exists(this.settingsFilePath))
                {
                    string pathsToSearch =
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Microsoft SQL Server\100\Tools\Shell\;");

                    pathsToSearch += ConfigurationManager.AppSettings["pathsToSearch"];
                    foreach (var path in pathsToSearch.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        string fileName = Path.Combine(Environment.ExpandEnvironmentVariables(path.Trim()), "SqlStudio.bin");
                        if (File.Exists(fileName))
                        {
                            this.settingsFilePath = fileName;
                            break;
                        }
                    }
                }

                return this.settingsFilePath;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Gets a unique name for the connection.
        /// </summary>
        /// <param name="serverConnectionItem">The server connection item.</param>
        /// <param name="existingServers">The existing servers.</param>
        /// <returns>A unique name for the connection</returns>
        private static string GetServerName(ServerConnectionItem serverConnectionItem, Dictionary<string, Server> existingServers)
        {
            var key = string.Format("{0} ({1})", serverConnectionItem.Instance, serverConnectionItem.AuthenticationMethod == 0 ? "Windows" : "Sql");
            var testKey = key;
            int i = 2;
            while (existingServers.Any(s => s.Value.Name == testKey))
            {
                testKey = key + " " + i;
                i++;
                if (existingServers.All(s => s.Value.Name != testKey))
                {
                    key = testKey;
                }
            }

            return key;
        }

        /// <summary>
        /// Handles the Click event of the CancelButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Refresh();
        }

        /// <summary>
        /// Handles the Click event of the DeleteLoginButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void DeleteLoginButton_Click(object sender, RoutedEventArgs e)
        {
            var serverPair = this.Servers.FirstOrDefault(s => s.Value.Name == this.uxServers.SelectedItem.ToString());
            foreach (var pair in this.sqlStudioSettings.SSMS.ConnectionOptions.ServerTypes)
            {
                ServerTypeItem serverTypeItem = pair.Value;
                var toRemove = serverTypeItem.Servers.Where(server =>
                    server.Instance == serverPair.Value.ServerConnectionItem.Instance)
                    .ToList();

                foreach (ServerConnectionItem serverConnectionItem in toRemove)
                {
                    var loginsToDelete =
                        serverConnectionItem.Connections.Where(c => c.UserName == this.uxLogins.SelectedItem.ToString())
                        .ToList();
                    foreach (var login in loginsToDelete)
                    {
                        serverConnectionItem.Connections.RemoveItem(login);
                    }
                }
            }

            this.HasChanges = true;
            this.GetLoginsForSelectedServer();
        }

        /// <summary>
        /// Handles the Click event of the DeleteServerButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void DeleteServerButton_Click(object sender, RoutedEventArgs e)
        {
            var serverPair = this.Servers.FirstOrDefault(s => s.Value.Name == this.uxServers.SelectedItem.ToString());

            foreach (var pair in this.sqlStudioSettings.SSMS.ConnectionOptions.ServerTypes)
            {
                ServerTypeItem serverTypeItem = pair.Value;
                var toRemove = serverTypeItem.Servers.Where(server =>
                    server.Instance == serverPair.Value.ServerConnectionItem.Instance)
                    .ToList();

                foreach (ServerConnectionItem serverConnectionItem in toRemove)
                {
                    serverTypeItem.Servers.RemoveItem(serverConnectionItem);
                }
            }

            this.Servers.Remove(serverPair.Key);

            this.HasChanges = true;
            this.Refresh(false);
        }

        /// <summary>
        /// Gets the logins for selected server.
        /// </summary>
        private void GetLoginsForSelectedServer()
        {
            this.uxDeleteLogin.IsEnabled = false;
            this.uxDeleteServer.IsEnabled = false;
            this.uxLogins.Items.Clear();
            if (this.uxServers.SelectedItem != null)
            {
                var serverPair = this.Servers.FirstOrDefault(s => s.Value.Name == this.uxServers.SelectedItem.ToString());
                if (serverPair.Value != null)
                {
                    this.uxDeleteServer.IsEnabled = true;
                    ServerConnectionItem serverInstance = serverPair.Value.ServerConnectionItem;
                    if (serverInstance.Connections.Any())
                    {
                        foreach (var connection in serverInstance.Connections.OrderBy(k => k.UserName))
                        {
                            this.uxLogins.Items.Add(connection.UserName);
                        }

                        this.uxLogins.SelectedIndex = 0;
                        this.uxDeleteLogin.IsEnabled = true;
                    }
                }
            }
        }

        /// <summary>
        /// Refreshes this instance.
        /// </summary>
        /// <param name="readFromFile">if set to <c>true</c> the settings are (re)read from file; otherwise the in-memory version is used.</param>
        private void Refresh(bool readFromFile = true)
        {
            if (File.Exists(this.SettingsFilePath))
            {
                if (readFromFile)
                {
                    this.servers = null;
                }

                int index = this.uxServers.SelectedIndex;
                if (index == -1)
                {
                    index = 0;
                }
                else if (index == this.Servers.Count)
                {
                    index--;
                }

                this.uxServers.Items.Clear();
                if (this.Servers.Any())
                {
                    foreach (var serverInstanceName in this.Servers.Keys.OrderBy(k => k))
                    {
                        this.uxServers.Items.Add(serverInstanceName);
                    }

                    this.uxServers.SelectedIndex = index;
                }
            }
            else
            {
                MessageBox.Show(
                    this,
                    "Unable to find SqlStudio.bin file. Manually find file and put path in config file.",
                    "File not found.",
                    MessageBoxButton.OK);
            }
        }

        /// <summary>
        /// Handles the Click event of the SaveButton control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            string backupPath = string.Format(
                "{0}_{1}.bin",
                Path.GetFileNameWithoutExtension(this.SettingsFilePath),
                DateTime.Now.ToString("yyyy_MM_dd_HHmmss"));
            string folder = Path.GetDirectoryName(this.settingsFilePath);
            if (folder != null)
            {
                backupPath = Path.Combine(folder, backupPath);
                File.Copy(this.SettingsFilePath, backupPath);
                var binaryFormatter = new BinaryFormatter();
                var outStream = new MemoryStream();
                binaryFormatter.Serialize(outStream, this.sqlStudioSettings);
                var outBytes = new byte[outStream.Length];
                outStream.Position = 0;
                outStream.Read(outBytes, 0, outBytes.Length);
                File.WriteAllBytes(this.SettingsFilePath, outBytes);
                this.HasChanges = false;
                this.Refresh();
            }
        }

        /// <summary>
        /// Handles the KeyDown event of the Servers control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Input.KeyEventArgs"/> instance containing the event data.</param>
        private void Servers_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F5)
            {
                this.Refresh();
            }
        }

        /// <summary>
        /// Handles the SelectionChanged event of the Servers control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.Controls.SelectionChangedEventArgs"/> instance containing the event data.</param>
        private void Servers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            this.GetLoginsForSelectedServer();
        }

        /// <summary>
        /// Handles the Loaded event of the Window control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.Windows.RoutedEventArgs"/> instance containing the event data.</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.Refresh();
            var assemblyName = Assembly.GetExecutingAssembly().GetName();
            this.Title = string.Format("{0} v{1}", assemblyName.Name, assemblyName.Version);
        }

        #endregion Methods

        #region Nested Types

        /// <summary>
        /// Class encapsulating the server
        /// </summary>
        private class Server
        {
            #region Constructors

            /// <summary>
            /// Initializes a new instance of the <see cref="Server"/> class.
            /// </summary>
            /// <param name="serverConnectionItem">The server connection item.</param>
            /// <param name="name">The name.</param>
            internal Server(ServerConnectionItem serverConnectionItem, string name)
            {
                this.ServerConnectionItem = serverConnectionItem;
                this.Name = name;
            }

            #endregion Constructors

            #region Properties

            /// <summary>
            /// Gets the name.
            /// </summary>
            internal string Name { get; private set; }

            /// <summary>
            /// Gets the server connection item.
            /// </summary>
            internal ServerConnectionItem ServerConnectionItem { get; private set; }

            #endregion Properties
        }

        #endregion Nested Types
    }
}