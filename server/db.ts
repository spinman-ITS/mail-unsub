export type UnsubscribedSender = {
  userEmail: string;
  senderAddress: string;
  senderDomain: string;
  method: string;
  createdAt: string;
};

type SenderInput = Omit<UnsubscribedSender, "createdAt">;

interface SenderStore {
  record(input: SenderInput): Promise<void>;
  listForUser(userEmail: string): Promise<UnsubscribedSender[]>;
}

class MemorySenderStore implements SenderStore {
  private rows: UnsubscribedSender[] = [];

  async record(input: SenderInput): Promise<void> {
    const key = `${input.userEmail}|${input.senderAddress}`;
    if (this.rows.some((row) => `${row.userEmail}|${row.senderAddress}` === key)) {
      return;
    }
    this.rows.push({ ...input, createdAt: new Date().toISOString() });
  }

  async listForUser(userEmail: string): Promise<UnsubscribedSender[]> {
    return this.rows.filter((row) => row.userEmail === userEmail);
  }
}

class PostgresSenderStore implements SenderStore {
  private ready: Promise<unknown>;

  constructor(private pool: import("pg").Pool) {
    this.ready = pool.query(`
      CREATE TABLE IF NOT EXISTS unsubscribed_senders (
        id BIGSERIAL PRIMARY KEY,
        user_email TEXT NOT NULL,
        sender_address TEXT NOT NULL,
        sender_domain TEXT NOT NULL,
        method TEXT NOT NULL,
        created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
        UNIQUE (user_email, sender_address)
      )
    `);
  }

  async record(input: SenderInput): Promise<void> {
    await this.ready;
    await this.pool.query(
      `INSERT INTO unsubscribed_senders (user_email, sender_address, sender_domain, method)
       VALUES ($1, $2, $3, $4)
       ON CONFLICT (user_email, sender_address) DO NOTHING`,
      [input.userEmail, input.senderAddress, input.senderDomain, input.method]
    );
  }

  async listForUser(userEmail: string): Promise<UnsubscribedSender[]> {
    await this.ready;
    const result = await this.pool.query(
      `SELECT user_email, sender_address, sender_domain, method, created_at
       FROM unsubscribed_senders WHERE user_email = $1 ORDER BY created_at DESC`,
      [userEmail]
    );
    return result.rows.map((row) => ({
      userEmail: row.user_email,
      senderAddress: row.sender_address,
      senderDomain: row.sender_domain,
      method: row.method,
      createdAt: new Date(row.created_at).toISOString()
    }));
  }
}

export async function createSenderStore(): Promise<SenderStore> {
  const connectionString = process.env.DATABASE_URL;

  if (!connectionString) {
    console.log("[db] DATABASE_URL not set, using in-memory sender store");
    return new MemorySenderStore();
  }

  const { Pool } = await import("pg");
  const pool = new Pool({ connectionString });
  console.log("[db] using Postgres sender store");
  return new PostgresSenderStore(pool);
}

export function normalizeEmail(value: string): string {
  return value.trim().toLowerCase();
}

export function emailDomain(address: string): string {
  const at = address.lastIndexOf("@");
  return at >= 0 ? address.slice(at + 1).toLowerCase() : "";
}
