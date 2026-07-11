export type SerialOperationQueue = Readonly<{
  enqueue<T>(task: () => T | Promise<T>): Promise<T>;
}>;

export function createSerialOperationQueue(): SerialOperationQueue {
  let tail: Promise<void> | null = null;

  return {
    enqueue<T>(task: () => T | Promise<T>): Promise<T> {
      let result: Promise<T>;
      if (tail) {
        result = tail.then(task);
      } else {
        try {
          result = Promise.resolve(task());
        } catch (error) {
          result = Promise.reject(error);
        }
      }
      const settled = result.then(
        () => undefined,
        () => undefined
      );
      tail = settled;
      void settled.then(() => {
        if (tail === settled) {
          tail = null;
        }
      });
      return result;
    }
  };
}
