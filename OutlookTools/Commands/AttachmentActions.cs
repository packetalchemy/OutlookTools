using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookTools.Commands
{
    /// <summary>
    /// OutlookTools — Attachment actions:
    /// - Reply (with attachments)
    /// - Reply All (with attachments)
    /// - Forward WITHOUT attachments (preserving inline signature images)
    ///
    /// Pure Outlook Interop. NO file copy to disk. NO network. NO cache.
    /// </summary>
    public static class AttachmentActions
    {
        /// <summary>
        /// Reply to the selected mail item.
        /// </summary>
        public static void Reply(bool keepAttachments, bool replyAll)
        {
            Outlook.MailItem original = GetCurrentMailItem();
            if (original == null) return;

            try
            {
                Outlook.MailItem reply = replyAll
                    ? (Outlook.MailItem)original.ReplyAll()
                    : (Outlook.MailItem)original.Reply();

                if (replyAll && !keepAttachments)
                {
                    // Edge case: user explicitly chose Reply All without attachments
                    StripFileAttachments(reply);
                }
                else if (!keepAttachments)
                {
                    StripFileAttachments(reply);
                }
                // else: keep attachments as-is (Outlook's Reply() already copies them)

                reply.Display(false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Reply failed: {ex.Message}", "OutlookTools",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
            finally
            {
                if (original != null) Marshal.ReleaseComObjectSafe(original);
            }
        }

        /// <summary>
        /// Forward the current mail item WITHOUT file attachments.
        /// Preserves inline signature images (CID references).
        /// </summary>
        public static void ForwardWithoutAttachments()
        {
            Outlook.MailItem original = GetCurrentMailItem();
            if (original == null) return;

            try
            {
                Outlook.MailItem forward = (Outlook.MailItem)original.Forward();
                StripFileAttachments(forward);
                forward.Display(false);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Forward failed: {ex.Message}", "OutlookTools",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
            finally
            {
                if (original != null) Marshal.ReleaseComObjectSafe(original);
            }
        }

        /// <summary>
        /// Remove file attachments from the mail item.
        /// Inline images (linked via Content-Id) are preserved.
        /// </summary>
        private static void StripFileAttachments(Outlook.MailItem item)
        {
            try
            {
                Outlook.Attachments attachments = item.Attachments;
                // Iterate from end → safe removal
                for (int i = attachments.Count; i >= 1; i--)
                {
                    var attachment = attachments[i];
                    try
                    {
                        // Preserve inline images — they have a Position-related property
                        // For file attachments: index is unique per item
                        // Inline images are typically embedded in HTMLBody
                        var position = SafeGetProperty<int>(attachment, "Position");
                        // If position > 0 the attachment is rendered in body — skip
                        if (position > 0)
                        {
                            // Inline image — keep
                            continue;
                        }
                        // Plain file attachment — remove
                        attachment.Delete();
                    }
                    finally
                    {
                        if (attachment != null) Marshal.ReleaseComObjectSafe(attachment);
                    }
                }

                // Best-effort: also remove tables of attachments in HTMLBody
                // We only remove attached file containers, NOT inline images.
                string html = item.HTMLBody ?? "";
                // Remove <img src="cid:..."> tags (inline) — DO NOT TOUCH
                // Remove file attachment section: typically Outlook inserts attachments
                // as links at end of body. We do NOT modify HTMLBody to avoid breaking
                // signature images. The file attachments were already deleted above.

                // Important: After deletion, refresh item
                item.Save();
            }
            catch (Exception ex)
            {
                ThisAddIn.LogDebug("StripFileAttachments: " + ex.Message);
            }
        }

        /// <summary>
        /// Read the currently-selected email in the Explorer.
        /// Returns null if nothing is selected or it's not a MailItem.
        /// </summary>
        private static Outlook.MailItem GetCurrentMailItem()
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null) return null;

                var selection = explorer.Selection;
                if (selection == null || selection.Count == 0) return null;

                var first = selection[1];
                if (first is Outlook.MailItem mail)
                    return mail;

                Marshal.ReleaseComObjectSafe(first);
                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Safe reflection helper for COM properties.
        /// </summary>
        private static T SafeGetProperty<T>(object comObject, string propName)
        {
            try
            {
                var val = comObject.GetType().InvokeMember(
                    propName,
                    System.Reflection.BindingFlags.GetProperty |
                    System.Reflection.BindingFlags.IgnoreReturn |
                    System.Reflection.BindingFlags.IgnoreCase,
                    null, comObject, null);
                return (T)Convert.ChangeType(val, typeof(T));
            }
            catch
            {
                return default(T);
            }
        }
    }

    /// <summary>
    /// Safe COM object release for Outlook interop.
    /// </summary>
    internal static class Marshal
    {
        public static void ReleaseComObjectSafe(object obj)
        {
            try
            {
                if (obj != null && System.Runtime.InteropServices.Marshal.IsComObject(obj))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
