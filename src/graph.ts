const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export async function graphRequest(
  accessToken: string,
  path: string,
  options: RequestInit = {}
): Promise<unknown> {
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    },
  });

  // DELETE returns 204, sendMail/reply return 202 — both have empty bodies
  if (response.status === 204 || response.status === 202) return null;

  const data = await response.json() as Record<string, unknown>;

  if (!response.ok) {
    const err = data.error as { message?: string } | undefined;
    throw new Error(`Graph API error (${response.status}): ${err?.message ?? response.statusText}`);
  }

  return data;
}

// Raw-response variant — used when we need binary content (e.g. file downloads)
// instead of JSON. Caller is responsible for consuming the body.
export async function graphRequestRaw(
  accessToken: string,
  path: string,
  options: RequestInit = {}
): Promise<Response> {
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...options.headers,
    },
    redirect: 'follow',
  });

  if (!response.ok) {
    let errMsg = response.statusText;
    try {
      const data = await response.json() as { error?: { message?: string } };
      errMsg = data.error?.message ?? errMsg;
    } catch { /* not JSON */ }
    throw new Error(`Graph API error (${response.status}): ${errMsg}`);
  }

  return response;
}
