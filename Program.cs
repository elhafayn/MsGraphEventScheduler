using CommandLine;
using CommandLine.Text;
using Microsoft.Graph.Models;
using System.Collections;
using System.Collections.Generic;

namespace MSGraph_Hack_Togother;
class Program
{
    public class Options
    {
        [Option('l', "list-users", Required = false, HelpText = "list all the users in your organization.")]
        public bool listUsers { get; set; }
        [Option('s', "search", Required = false, HelpText = "search for users by email")]
        public string keyword { get; set; }
        [Option('t', "top", Required = false, HelpText = "limit the resource count")]
        public int top { get; set; }

        [Option('q', "quit", Required = false, HelpText = "close the application")]
        public bool quit { get; set; }

        #region findMeetingTimes
        [Option('r', "attendees", Required = false, HelpText = "A collection of attendees or resources for the meeting")]
        public IEnumerable<string> attendees { get; set; }

        [Option( "la", Required = false, HelpText = "list selected attendees")]
        public bool listAttendees { get; set; }
        [Option("isOrganizerOptional", Required = false, HelpText = "Specify True if the organizer doesn't necessarily have to attend")]
        public bool isOrganizerOptional { get; set; }
        [Option('a', "add-attendee", Required = false, HelpText = "Add recipient  email to the meeting")]
        public string attendee { get; set; }
        [Option('f', "findMeetingTimes", Required = false, HelpText = "Suggest meeting times ")]
        public bool findMeetingTimes { get; set; }
        [Option("meetingDuration", Required = false, HelpText = "The length of the meeting in minutes ")]
        public long meetingDuration { get; set; }
        [Option("map", Required = false, HelpText = "The minimum required confidence for a time slot to be returned in the response. It is a % value ranging from 0 to 100.")]
        public long minimumAttendeePercentage { get; set; }
        [Option("timeConstraint", Required = false, HelpText = "list slots of point of time in a combined date and time  representation ({date}T{time}; for example, 2023-04-16T09:00:00  separated by space, Note time zone is Pacific Standard Time")]
        public IEnumerable<string> timeConstraint { get; set; }
        #endregion
        [Option("schedule", Required = false, HelpText = "suggested time key to schedule the meeting ")]
        public int schedule { get; set; }
        [Option("sub", Required = false, HelpText = " meeting subject ")]
        public string subject { get; set; }
        [Option("con", Required = false, HelpText = " meeting content ")]
        public string content { get; set; }  
        
        [Option("alt", Required = false, HelpText = " Meeting Allow New Time Proposals ")]
        public bool AllowNewTimeProposals { get; set; }

    }
    static SechdulerService service = new SechdulerService();
    static Dictionary<int, SuggestedMeetingTimes> _suggestedMeetingTimes = new Dictionary<int, SuggestedMeetingTimes>();
    static void Main(string[] args)
    {
        MainAsync().Wait();

    }
    static async Task MainAsync()
    {
        var settings = Settings.LoadSettings();
        // // Initialize Graph

        service.InitializeGraphForUserAuth(settings, (info, cancel) =>
        {
            // Display the device code message to
            // the user. This tells them
            // where to go to sign in and provides the
            // code to use.
            Console.WriteLine(info.Message);
            return Task.FromResult(0);
        });

        var AuthUser = await service.GetAuthorizedUserAsync();

        Console.WriteLine($"Hello, {AuthUser?.DisplayName}!");
        // For Work/school accounts, email is in Mail property
        // Personal accounts, email is in UserPrincipalName
        Console.WriteLine($"Email: {AuthUser?.Mail ?? AuthUser?.UserPrincipalName ?? ""}");


        string choice = "--help";

        while (choice != null)
        {
            try
            {

                await StartUp(choice.Split(" "));
                choice = Console.ReadLine();

            }
            catch (System.FormatException ex)
            {
                // Set to invalid value
                Console.WriteLine(ex.Message);
                choice = null;
            }
        }
    }
    static Task StartUp(string[] args)
    {
        var parser = new CommandLine.Parser(with => with.HelpWriter = null);
        var parserResult = parser.ParseArguments<Options>(args);
        parserResult
         .WithParsed<Options>(async options => await Run(options))
         .WithNotParsed(errs => DisplayHelp(parserResult, errs));
        return Task.CompletedTask;
    }

    static Task DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
    {
        var helpText = HelpText.AutoBuild(result, h =>
         {
             h.AdditionalNewLineAfterOption = false; //remove the extra newline between options
             h.Heading = "Meeting Scheduler 1.0.0-beta"; //change header
             h.Copyright = "Copyright (c) 2023 "; //change copyrigt text
             return HelpText.DefaultParsingErrorsHandler(result, h);
         }, e => e);
        Console.WriteLine(helpText);
        return Task.CompletedTask;
    }
    static async Task Run(Options o)
    {
        if (o.quit)
        {
            //close application
            Environment.Exit(0);
        }
        else if (o.listUsers)
        {

            if (!string.IsNullOrEmpty(o.keyword))
                await SearchUsers(o.keyword, o.top);
            else
                await getUsers(o.top);
        }
        else if (!string.IsNullOrEmpty(o.attendee)) {
          
            var user = await service.GetUserByEmail(o.attendee);
            if (user != null)
            {
                service.addAttendee(new EmailAddress() { Address = user.Mail, Name = user.DisplayName }, AttendeeType.Required);
                Console.WriteLine("user has been added successfully to the meeting Attendees !");
            }
                

        }
        else if (o.listAttendees)
        {

            displaySelectedAttendess();
        }
        else if (o.attendees!=null&&o.attendees?.Count()>0)
        {

            await addAttendees(o.attendees);
        }   
        else if (o.findMeetingTimes)
        {

            await getSuggestedMeetingTimes(o.timeConstraint,o.meetingDuration,o.isOrganizerOptional);
        }
        else if (o.schedule>0)
        {
            if (string.IsNullOrEmpty(o.subject))
                Console.WriteLine("Please use --sub to enter the meeting subject");
            else if (string.IsNullOrEmpty(o.content))
                    Console.WriteLine("Please use --con to enter the meeting content");
            else
            await scheduleTheMeeting(o.schedule,o.subject, o.content, o.AllowNewTimeProposals);
        }


    }

    private static async Task scheduleTheMeeting(int meetingKey,string subject, string content, bool allowNewTimeProposals)
    {
        if (_suggestedMeetingTimes.TryGetValue(meetingKey, out SuggestedMeetingTimes suggestedMeeting))
        {
            var meeting = await service.CreateMeetingAsync(subject, content, suggestedMeeting.start, suggestedMeeting.end, allowNewTimeProposals);
            if (meeting != null)
                Console.WriteLine($"{meeting.WebLink}");

        }
        else {
            if (_suggestedMeetingTimes.Count > 0) {
                Console.WriteLine($"please choose one of the suggested times");
            }
            
        
        }
    }

    private static async  Task getSuggestedMeetingTimes(IEnumerable<string> timeConstraint,long meetingDuration,bool isOrganizerOptional)
    {
        var times = timeConstraint?.Count()>0?new TimeConstraint() { TimeSlots = timeConstraint.Select(x => new TimeSlot() { Start = new DateTimeTimeZone { DateTime = x, TimeZone = "Pacific Standard Time" } }).ToList() }:null;
        var suggestedTimes = await service.FindMeetingTimes(times, isOrganizerOptional, TimeSpan.FromMinutes(meetingDuration), true, 50);
        if (suggestedTimes != null)
        {
            _suggestedMeetingTimes.Clear();
            int index = 1;
            foreach (var suggestedTime in suggestedTimes?.MeetingTimeSuggestions)
            {
                Console.WriteLine($"{suggestedTime.MeetingTimeSlot.Start.DateTime} - {suggestedTime.MeetingTimeSlot.End.DateTime}, {suggestedTime.Confidence}");
                _suggestedMeetingTimes.Add(index, new SuggestedMeetingTimes() { start = suggestedTime.MeetingTimeSlot.Start, end = suggestedTime.MeetingTimeSlot.End });
                index++;
            }
            Console.WriteLine($"{suggestedTimes.EmptySuggestionsReason}");

        }

   
    }

    private static async Task addAttendees(IEnumerable<string> attendees)
    {
        foreach (var attendee in attendees) {
        
            var user = await service.GetUserByEmail(attendee);
            if (user != null)
            {
                service.addAttendee(new EmailAddress() { Address = user.Mail, Name = user.DisplayName }, AttendeeType.Required);

            }
            else {
                Console.WriteLine($"user {attendee} not found, please check the entered name is correct and users are separated by space ");
            }
        }
       
    }

    private static void displaySelectedAttendess()
    {
       var attendees=service.getSelectedAttendees();

        foreach (var attendee in attendees) {
            Console.WriteLine($"Mail: {attendee.EmailAddress.Address}, Type: {attendee.Type.ToString()} ");
        }
    }

    static async Task SearchUsers(string keyword, int top=10)
    {
        try
        {
            var users = await service.SearchUsers(keyword,top);
            foreach (var u in users)
            {
                Console.WriteLine($"{u.DisplayName} - {u.Mail}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting user: {ex.Message}");
        }
    }

    static async Task getUsers(int limit=10)
    {
        try
        {
            var users = await service.ListUsers(limit);
            foreach (var u in users)
            {
                Console.WriteLine($"{u.DisplayName} - {u.Mail}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting user: {ex.Message}");
        }
    }





}
