using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.TermStore;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Http.Generated;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.Graph.Constants;

namespace MSGraph_Hack_Togother
{
    public class SechdulerService:ISechdulerService
    {
        // Settings object
        private  Settings? _settings;
        // User auth token credential
        private  DeviceCodeCredential? _deviceCodeCredential;
        // Client configured with user authentication
        private  GraphServiceClient? _userClient;

        private HashSet<AttendeeBase> _attendees=Enumerable.Empty<AttendeeBase>().ToHashSet();

        private HashSet<User> _users=Enumerable.Empty<User>().ToHashSet();

        private  async Task<string> GetUserTokenAsync()
        {
            // Ensure credential isn't null
            _ = _deviceCodeCredential ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            // Ensure scopes isn't null
            _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

            // Request token with given scopes
            var context = new TokenRequestContext(_settings.GraphUserScopes);
            var response = await _deviceCodeCredential.GetTokenAsync(context);
            return response.Token;
        }

        public void addAttendee(EmailAddress attendeeEmail, AttendeeType type)
        {
            var attendeeObj = _attendees.SingleOrDefault(x => x.EmailAddress.Equals(attendeeEmail));
            if (_attendees.TryGetValue(attendeeObj, out AttendeeBase exist))
            {
                exist.Type = type;
                _attendees.Remove(attendeeObj);
                _attendees.Add(exist);
            }
            else {
                _attendees.Add(new AttendeeBase() {EmailAddress=attendeeEmail,Type=type });

            }
        }

        public void addUsersToCache(User user)
        {
            _users.Add(user);
        }

        public async Task<Event> CreateMeetingAsync(string subject, string content, DateTimeTimeZone start, DateTimeTimeZone end, bool AllowNewTimeProposals)
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

           

            var requestBody = new Event
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = content,
                },
                Start =start,
                End = end,

                Attendees = _attendees.Select(a => new Attendee
                {
                    EmailAddress = a.EmailAddress,
                    Type = a.Type,
                }).ToList(),
                IsOnlineMeeting = true,
                AllowNewTimeProposals=AllowNewTimeProposals,
                OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness,
                TransactionId = Guid.NewGuid().ToString(),
            };
            var result = await _userClient.Me.Events.PostAsync(requestBody, (requestConfiguration) =>
            {
                requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
            });

            return result;

        }

        public async Task<MeetingTimeSuggestionsResult> FindMeetingTimes( TimeConstraint timeConstraint, bool IsOrganizerOptional, TimeSpan MeetingDuration, bool ReturnSuggestionReasons, double MinimumAttendeePercentage)
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var requestBody = new Microsoft.Graph.Me.FindMeetingTimes.FindMeetingTimesPostRequestBody
            {
                Attendees = _attendees.ToList(),
             
                TimeConstraint = timeConstraint,
                IsOrganizerOptional = IsOrganizerOptional,
                MeetingDuration = MeetingDuration,
                ReturnSuggestionReasons = ReturnSuggestionReasons,
                MinimumAttendeePercentage = MinimumAttendeePercentage,
            };
            try
            {
                var result = await _userClient.Me.FindMeetingTimes.PostAsync(requestBody, (requestConfiguration) =>
                {
                    requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Pacific Standard Time\"");
                });
                return result;
            }
            catch (Exception ex) {

                return null;
            }
          
        }

        public async Task<User> GetAuthorizedUserAsync()
        {
            // Ensure client isn't null
            _ = _userClient ??
                throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var req = await _userClient.Me.GetAsync();

            return new User
            {
                // Only request specific properties
                DisplayName = req.DisplayName,
                Mail = req.Mail,
                UserPrincipalName = req.UserPrincipalName
            };
        }

        public void InitializeGraphForUserAuth(Settings settings, Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
        {
            _settings = settings;

            _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
                settings.TenantId, settings.ClientId);

            _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
        }

        public async Task<IEnumerable<User>> ListUsers(int? limit)
        {
            if(_users.Count>limit)
                return _users.Take(limit??10);

            _ = _userClient ??
       throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var result = await _userClient.Users.GetAsync((requestConfiguration) => {
                requestConfiguration.QueryParameters.Top = limit>0?limit:10;
            });
            return result?.Value;
        }

        public void removeAttendee(EmailAddress attendeeEmail)
        {
           _attendees.RemoveWhere(x => x.EmailAddress.Equals(attendeeEmail));
        }

        public async Task<IEnumerable<User>> SearchUsers(string? keyword, int? limit)
        {
            _ = _userClient ??
  throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var result = await _userClient.Users.GetAsync((requestConfiguration) => {
                requestConfiguration.QueryParameters.Top = limit > 0 ? limit : 10;
                requestConfiguration.QueryParameters.Search = $"\"mail:{keyword}\"";
                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
            });
            return result?.Value;
        }
        public List<AttendeeBase> getSelectedAttendees() {
            return _attendees.ToList();
        }

        public async Task<User> GetUserByEmail(string email)
        {
            _ = _userClient ??
throw new System.NullReferenceException("Graph has not been initialized for user auth");

            var result = await _userClient.Users.GetAsync((requestConfiguration) => {
                requestConfiguration.QueryParameters.Count = true;
                requestConfiguration.Headers.Add("ConsistencyLevel", "eventual");
                requestConfiguration.QueryParameters.Filter = $"mail eq '{email}'";
            });
            return result?.Value.SingleOrDefault();
        }
    }
}
