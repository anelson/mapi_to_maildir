using System;
using System.Windows.Forms;
using MAPI33;

namespace MapiToMaildir {
	/// <summary>
	/// A tree node corresponding to a MAPI folder
	/// </summary>
	internal class MapiFolderNode : TreeNode, IDisposable {
		IMAPIFolder _folder;

		public MapiFolderNode(IMAPIFolder folder) {
			Text = MapiUtils.GetStringProperty(folder, Tags.PR_DISPLAY_NAME);
			_folder = folder.Clone();
		}

		public IMAPIFolder Folder {
			get {
				return _folder;
			}
		}

		#region IDisposable Members

		public void Dispose() {
			//Have to clean up the folder object
            if (_folder != null) {
                _folder.Dispose();
                _folder = null;
            }
		}

		#endregion
	}
}
