using Azure.Identity;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MSGraph_Hack_Togother
{
    public interface ISechdulerService
    {
        Task<IEnumerable<User>> ListUsers(int? limit);

        Task<IEnumerable<User>> SearchUsers(string? keyword,int? limit);
        Task<User> GetUserByEmail(string email);

        void InitializeGraphForUserAuth(Settings settings,
    Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt);

        Task<User> GetAuthorizedUserAsync();


        Task<MeetingTimeSuggestionsResult> FindMeetingTimes(
            TimeConstraint timeConstraint,bool IsOrganizerOptional, TimeSpan MeetingDuration, bool ReturnSuggestionReasons,double MinimumAttendeePercentage);

        Task<Event> CreateMeetingAsync(string subject,string content, DateTimeTimeZone start, DateTimeTimeZone end,bool AllowNewTimeProposals);


        public void addAttendee(EmailAddress attendeeEmail, AttendeeType type);
        public void removeAttendee(EmailAddress attendeeEmail);

        public void addUsersToCache(User user);

    }
}
