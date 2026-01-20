using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;

namespace ActionableMessageSample.Services;

/// <summary>
/// Default in-memory implementation of <see cref="IActionableMessageStateStore"/>.
/// </summary>
public sealed class InMemoryActionableMessageStateStore : IActionableMessageStateStore
{
    private readonly ConcurrentDictionary<string, ConcurrentDictionary<string, byte>> _completedActions = new(StringComparer.OrdinalIgnoreCase);

    /// <inheritdoc />
    public bool TryMarkActionCompleted(string messageKey, string actionName)
    {
        if (string.IsNullOrWhiteSpace(messageKey) || string.IsNullOrWhiteSpace(actionName))
        {
            return false;
        }

        var actionsForMessage = _completedActions.GetOrAdd(
            messageKey,
            static _ => new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase));

        return actionsForMessage.TryAdd(actionName, 0);
    }

    /// <inheritdoc />
    public IReadOnlyCollection<string> GetCompletedActions(string messageKey)
    {
        if (string.IsNullOrWhiteSpace(messageKey))
        {
            return Array.Empty<string>();
        }

        if (_completedActions.TryGetValue(messageKey, out var actions))
        {
            return actions.Keys.ToArray();
        }

        return Array.Empty<string>();
    }
}
