using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using MAPI33;
using MAPI33.MapiTypes;
using MapiGuts = __MAPI33__INTERNALS__;

namespace MapiToMaildir
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class MainFrm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button cmdConnect;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtExportPath;
		private System.Windows.Forms.Button cmdBrowse;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.TreeView treeFolders;
		private System.Windows.Forms.Button cmdExport;
		private System.Windows.Forms.Label lblStatus;

		IMAPISession _session;

		public MainFrm()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );

            DisposeMapiStuff();
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.cmdConnect = new System.Windows.Forms.Button();
			this.treeFolders = new System.Windows.Forms.TreeView();
			this.label1 = new System.Windows.Forms.Label();
			this.txtExportPath = new System.Windows.Forms.TextBox();
			this.cmdBrowse = new System.Windows.Forms.Button();
			this.cmdExport = new System.Windows.Forms.Button();
			this.lblStatus = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// cmdConnect
			// 
			this.cmdConnect.Location = new System.Drawing.Point(16, 16);
			this.cmdConnect.Name = "cmdConnect";
			this.cmdConnect.Size = new System.Drawing.Size(128, 23);
			this.cmdConnect.TabIndex = 0;
			this.cmdConnect.Text = "Connect to Outlook";
			this.cmdConnect.Click += new System.EventHandler(this.cmdConnect_Click);
			// 
			// treeFolders
			// 
			this.treeFolders.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.treeFolders.CheckBoxes = true;
			this.treeFolders.ImageIndex = -1;
			this.treeFolders.Location = new System.Drawing.Point(16, 56);
			this.treeFolders.Name = "treeFolders";
			this.treeFolders.SelectedImageIndex = -1;
			this.treeFolders.Size = new System.Drawing.Size(656, 312);
			this.treeFolders.Sorted = true;
			this.treeFolders.TabIndex = 1;
			this.treeFolders.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeFolders_AfterCheck);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label1.Location = new System.Drawing.Point(24, 376);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 23);
			this.label1.TabIndex = 2;
			this.label1.Text = "Export to:";
			// 
			// txtExportPath
			// 
			this.txtExportPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtExportPath.Location = new System.Drawing.Point(96, 376);
			this.txtExportPath.Name = "txtExportPath";
			this.txtExportPath.Size = new System.Drawing.Size(528, 20);
			this.txtExportPath.TabIndex = 3;
			this.txtExportPath.Text = "";
			// 
			// cmdBrowse
			// 
			this.cmdBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmdBrowse.Location = new System.Drawing.Point(632, 376);
			this.cmdBrowse.Name = "cmdBrowse";
			this.cmdBrowse.Size = new System.Drawing.Size(40, 23);
			this.cmdBrowse.TabIndex = 4;
			this.cmdBrowse.Text = "...";
			// 
			// cmdExport
			// 
			this.cmdExport.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmdExport.Location = new System.Drawing.Point(312, 448);
			this.cmdExport.Name = "cmdExport";
			this.cmdExport.TabIndex = 5;
			this.cmdExport.Text = "Export";
			this.cmdExport.Click += new System.EventHandler(this.cmdExport_Click);
			// 
			// lblStatus
			// 
			this.lblStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblStatus.Location = new System.Drawing.Point(24, 416);
			this.lblStatus.Name = "lblStatus";
			this.lblStatus.Size = new System.Drawing.Size(648, 23);
			this.lblStatus.TabIndex = 6;
			// 
			// MainFrm
			// 
			this.ClientSize = new System.Drawing.Size(688, 478);
			this.Controls.Add(this.lblStatus);
			this.Controls.Add(this.cmdExport);
			this.Controls.Add(this.cmdBrowse);
			this.Controls.Add(this.txtExportPath);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.treeFolders);
			this.Controls.Add(this.cmdConnect);
			this.Name = "MainFrm";
			this.Text = "Adam Nelson\'s MAPI to Maildir";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new MainFrm());
		}

		protected override void OnLoad(EventArgs e) {
			base.OnLoad (e);
			MAPI33.MAPI.Initialize(null);
		}

		protected override void OnClosed(EventArgs e) {
			//Disconnect
			DisposeMapiStuff();

			MAPI33.MAPI.Uninitialize();
			base.OnClosed (e);
		}

		private void DisposeMapiStuff() {
            foreach (TreeNode node in treeFolders.Nodes) {
                DisposeMapiNode(node);
            }

			if (_session != null) {
				_session.Logoff(this.Handle, 0);
				_session.Dispose();
				_session = null;
			}
		}

        private void DisposeMapiNode(TreeNode node) {
            foreach (TreeNode childNode in node.Nodes) {
                DisposeMapiNode(childNode);
            }

            node.Remove();
            if (node is IDisposable) {
                ((IDisposable)node).Dispose();
            }
        }



		private void cmdConnect_Click(object sender, System.EventArgs e)
		{
			//Connect to the default outlook session
			if (_session != null) {
				_session.Dispose();
			}

			treeFolders.Nodes.Clear();

			//'log on', using the default profile
			MAPI33.MAPI.LogonEx(this.Handle, null, null, MAPI.FLAGS.UseDefaultProfile, out _session);

			//For each store, attempt to enumerate the folders therein
			IMAPITable stores;
			Error hr;

			//God, this is just like C++.  What a shit wrapper on MAPI!
			hr = _session.GetMsgStoresTable(0, out stores);
			if (hr != Error.Success) {
				throw new MapiException(hr);
			}

			using (stores) {
				Value[,] rows;
				hr = stores.QueryRows(100, 0, out rows);
				if (hr != Error.Success) {
					throw new MapiException(hr);
				}

				//Somehow, we know that the first element of the second rank is the store ID
				for (int idx = 0; idx < rows.GetLength(0); idx++) {
					IMsgStore store;

					hr = _session.OpenMsgStore(IntPtr.Zero, ENTRYID.Create(rows[idx, 0]), Guid.Empty, IMAPISession.FLAGS.BestAccess, out store);

					if (hr != Error.Success) {
						//Report this failure but otherwise don't worry about it
						MapiException ex = new MapiException(hr);

						MessageBox.Show(ex.Message);

						continue;
					}

					using (store) {
						try {
							ShowStoreInTree(store);
						} catch (Exception ex) {
							//Don't let a failure w/ one store stop the processing of other stores
							MessageBox.Show(ex.Message);
						}
					}
				}
			}
		}

		private void ShowStoreInTree(IMsgStore store) {
			//Add the stuff in this store to the tree
			MapiStoreNode storeNode = new MapiStoreNode(store);
			treeFolders.Nodes.Add(storeNode);

			//Get the ID of the root folder
			ENTRYID rootIpmFolder = MapiUtils.GetEntryIdProperty(store, Tags.PR_IPM_SUBTREE_ENTRYID);

			MAPI33.MAPI.TYPE objType;
			IUnknown unk;
			Error hr = store.OpenEntry(rootIpmFolder, Guid.Empty, 0, out objType, out unk);
			if (hr != Error.Success) {
				throw new MapiException(hr);
			}

			using (unk) {
				IMAPIFolder folder = (IMAPIFolder)unk;
				using (folder) {
					ShowFolderInTree(storeNode, folder);
				}
			}
		}

		private void ShowFolderInTree(TreeNode parentNode, IMAPIFolder folder) {
			MapiFolderNode folderNode = new MapiFolderNode(folder);
			parentNode.Nodes.Add(folderNode);
			
			//Enumerate the child folders
			IMAPITable childFoldersTable;
			Error hr = folder.GetHierarchyTable(0, out childFoldersTable);
			if (hr != Error.Success) {
				throw new MapiException(hr);
			}

			using (childFoldersTable) {
				Value[,] rows;
				//Retrieve only one column: the entry ID
				childFoldersTable.SetColumns(new Tags[] {Tags.PR_ENTRYID}, IMAPITable.FLAGS.Default);

				//Query rows one at a time
				for (; childFoldersTable.QueryRows(1, 0, out rows) == Error.Success && rows.Length > 0; ) {
					//Get the child folder having this row ID
					IUnknown unk;
					MAPI.TYPE objType;
					IMAPIFolder subFolder;

					hr = folder.OpenEntry(ENTRYID.Create(rows[0,0]), Guid.Empty, IMAPIContainer.FLAGS.BestAccess, out objType, out unk);
					if (hr != Error.Success) {
						throw new MapiException(hr);
					}

					subFolder = (IMAPIFolder)unk;
					unk.Dispose();

					using (subFolder) {
						ShowFolderInTree(folderNode, subFolder);
					}
				}
			}
		}

		private void cmdExport_Click(object sender, System.EventArgs e) {
			//Traverse all the nodes in the tree.  If any node is checked, export that node
			//Note that store nodes cannot store messages themselves, but they have a check box
			//to check all child folders
			try {
				Cursor.Current = Cursors.WaitCursor;

				String exportPath = txtExportPath.Text + @"\";

				foreach (TreeNode node in treeFolders.Nodes) {
					ProcessNode(exportPath, node);
				} 

				lblStatus.Text = "Done";
			} finally {
				this.Cursor = Cursors.Default;
			}
		}

		private void ProcessNode(String exportPath, TreeNode node) {
			//Process this node.  If it's checked, export its messages 
			//to a folder created for this node in exportPath.
			//Regardless of checked status, also process children

			//If this node is checked or has an ancetor who is, a maildir must be created
			//Only export messages if the node itself is a folder node, and is checked
			if (node.Checked || IsAncestorChecked(node)) {
				lblStatus.Text = String.Format("Processing folder {0}", node.Text);
				Application.DoEvents();

				MaildirWriter writer = MaildirWriter.CreateMailDir(exportPath, node.Text);

				if (node.Checked && node is MapiFolderNode) {
					writer.AddFolder(((MapiFolderNode)node).Folder);
				}

				//Process the child nodes
				foreach (TreeNode childNode in node.Nodes) {
					ProcessNode(writer.Path, childNode);
				}
			}
		}

		private bool IsAncestorChecked(TreeNode node) {
			//Checks if any ancestor of the node is checked.
			//If so, a Maildir must be created for this node even if its messages
			//aren't being exported
			foreach (TreeNode childNode in node.Nodes) {
				if (childNode.Checked || IsAncestorChecked(childNode)) {
					return true;
				}
			}

			return false;
		}

		private void treeFolders_AfterCheck(object sender, System.Windows.Forms.TreeViewEventArgs e) {
			//Check all children
			CheckChildren(e.Node, e.Node.Checked);
		}

		private void CheckChildren(TreeNode node, bool check) {
			foreach (TreeNode childNode in node.Nodes) {
				CheckChildren(childNode, check);
				childNode.Checked = check;
			}
		}
	}
}
