using System;
using System.Windows.Forms;
using MAPI33;

namespace MapiToMaildir
{
	/// <summary>
	/// A tree node corresponding to a MAPI message store
	/// </summary>
	internal class MapiStoreNode : TreeNode, IDisposable
	{
		IMsgStore _store;

		public MapiStoreNode(IMsgStore store)
		{
			Text = MapiUtils.GetStringProperty(store, Tags.PR_DISPLAY_NAME);
			_store = store.Clone();
		}

		public IMsgStore Store {
			get {
				return _store;
			}
		}

		#region IDisposable Members

		public void Dispose() {
			//Have to clean up the store object
            if (_store != null) {
                _store.Dispose();
                _store = null;
            }
		}

		#endregion
	}
}
