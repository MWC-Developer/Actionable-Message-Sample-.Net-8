using System.Collections.Generic;

namespace ActionableMessageSample.Services;

/// <summary>
/// Stores actionable message state in memory so each action only runs once per message.
/// </summary>
public interface IActionableMessageStateStore
{
    /// <summary>
    /// Marks an action as completed for the specified message key.
    /// </summary>
    /// <param name="messageKey">Unique key representing the actionable message.</param>
    /// <param name="actionName">Action identifier that completed.</param>
    /// <returns><c>true</c> when the action was recorded for the first time; otherwise <c>false</c>.</returns>
    bool TryMarkActionCompleted(string messageKey, string actionName);

    /// <summary>
    /// Gets the set of actions already completed for the provided message key.
    /// </summary>
    /// <param name="messageKey">Unique key representing the actionable message.</param>
    /// <returns>Collection of normalized action names already executed.</returns>
    IReadOnlyCollection<string> GetCompletedActions(string messageKey);
}
