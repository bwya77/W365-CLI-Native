namespace W365Cli;

/// <summary>
/// Small helper for turning a serial "foreach item, await one Graph call" loop into a
/// bounded-concurrency parallel fetch. Several screens (all-snapshots, group name/member
/// resolution, user experience sync overview) previously awaited one Graph request per item in
/// a tight loop — for a fleet with dozens of Cloud PCs or policies, that's dozens of sequential
/// round-trips where a handful of concurrent ones would do, and it multiplies Graph throttling
/// risk since each serial call is its own throttling opportunity.
///
/// A bounded concurrency cap (rather than unbounded Task.WhenAll) avoids hammering Graph with
/// e.g. 200 simultaneous requests for a very large fleet, which would likely trigger throttling
/// immediately anyway.
/// </summary>
internal static class ConcurrencyHelper
{
    /// <summary>
    /// Projects <paramref name="source"/> through <paramref name="selector"/> with at most
    /// <paramref name="maxConcurrency"/> operations in flight at once. Results are returned in
    /// the same order as <paramref name="source"/>, regardless of completion order.
    /// </summary>
    public static async Task<TResult[]> MapWithConcurrencyAsync<TSource, TResult>(
        IEnumerable<TSource> source,
        int maxConcurrency,
        Func<TSource, Task<TResult>> selector)
    {
        using var throttle = new SemaphoreSlim(Math.Max(1, maxConcurrency));

        async Task<TResult> RunOneAsync(TSource item)
        {
            await throttle.WaitAsync();
            try
            {
                return await selector(item);
            }
            finally
            {
                throttle.Release();
            }
        }

        return await Task.WhenAll(source.Select(RunOneAsync));
    }
}
