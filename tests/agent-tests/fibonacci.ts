// Calculates the nth Fibonacci number using an iterative approach.
// The Fibonacci sequence is defined as:
// F(0) = 0, F(1) = 1, and F(n) = F(n-1) + F(n-2) for n >= 2.
//
// This implementation avoids recursion to prevent call stack overhead
// and uses a simple loop to build up results efficiently.

export function fibonacci(n: number): number {
  // Guard for non-positive input values: return 0 by definition.
  if (n <= 0) return 0;

  // Base case: Fibonacci(1) = 1.
  if (n === 1) return 1;

  // Initialize the first two Fibonacci numbers.
  let a = 0; // Represents F(0)
  let b = 1; // Represents F(1)

  // Iteratively compute Fibonacci numbers up to n.
  for (let i = 2; i <= n; i++) {
    const next = a + b; // Compute next Fibonacci number.
    a = b; // Shift the previous values forward.
    b = next; // Update current Fibonacci number.
  }

  // Final value of b is F(n).
  return b;
}
