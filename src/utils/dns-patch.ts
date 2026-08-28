import dns from "node:dns";

/**
 * Patch dns.lookup to gracefully fall back to c-ares resolve4 when
 * libc getaddrinfo fails with EBUSY or ENOTFOUND in constrained container environments.
 */
export function setupDnsLookupFallback(): void {
  const origLookup = dns.lookup;
  if (!origLookup) return;

  try {
    if (typeof dns.setDefaultResultOrder === "function") {
      dns.setDefaultResultOrder("ipv4first");
    }
  } catch {
    // Ignore if not supported
  }

  // @ts-expect-error - overriding dns.lookup
  dns.lookup = (
    hostname: string,
    options: any,
    callback?: any
  ) => {
    let cb: any;
    let opts: any = {};

    if (typeof options === "function") {
      cb = options;
    } else {
      if (typeof options === "number") {
        opts = { family: options };
      } else if (options) {
        opts = options;
      }
      cb = callback;
    }

    origLookup(hostname, opts, (err: any, address: any, family: any) => {
      if (err && (err.code === "EBUSY" || err.code === "ENOTFOUND")) {
        dns.promises.resolve4(hostname)
          .then((ips) => {
            if (ips && ips.length > 0) {
              if (opts && opts.all) {
                cb(null, ips.map((ip) => ({ address: ip, family: 4 })));
              } else {
                cb(null, ips[0], 4);
              }
              return;
            }
            cb(err, address, family);
          })
          .catch(() => {
            cb(err, address, family);
          });
        return;
      }
      cb(err, address, family);
    });
  };
}
