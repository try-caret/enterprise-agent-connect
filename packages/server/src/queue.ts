/** A queue worker for one process: `next` handles one item and says whether there was one.
 *  One drain at a time; a kick during one makes it look again before stopping. */
export function drainer(name: string, next: () => Promise<boolean>) {
  let draining: Promise<void> | null = null;
  let again = false;
  let stopped = false;

  async function drain() {
    while (again) {
      again = false;
      while (!stopped && (await next()));
    }
  }

  function kick() {
    if (stopped) return;
    again = true;
    draining ??= drain()
      .catch((error) => console.error(`[${name}] worker`, error))
      .finally(() => {
        draining = null;
        if (again) kick();
      });
  }

  return {
    kick,
    start() {
      stopped = false;
    },
    /** Stops claiming and waits for the item in hand, so a rollout neither strands it nor
     *  handles it twice. */
    async stop() {
      stopped = true;
      again = false;
      while (draining) await draining;
    },
    async idle() {
      while (draining) await draining;
    },
  };
}
