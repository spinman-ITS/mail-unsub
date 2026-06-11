export type CurrentMessageInfo = {
  userEmail: string;
  senderAddress: string;
  senderName: string;
};

export function readCurrentMessageInfo(): CurrentMessageInfo | null {
  const mailbox = globalThis.Office?.context?.mailbox;
  const from = mailbox?.item?.from;

  if (!mailbox?.userProfile?.emailAddress || !from?.emailAddress) {
    return null;
  }

  return {
    userEmail: mailbox.userProfile.emailAddress,
    senderAddress: from.emailAddress,
    senderName: from.displayName || from.emailAddress
  };
}
