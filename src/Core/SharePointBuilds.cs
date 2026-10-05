using System;

namespace SharePointExplorer
{
    public enum SharePointGeneration
    {
        Unknown = 0,
        SharePoint2013OrEarlier = 1,
        SharePoint2016 = 2,
        SharePoint2019 = 3,
        SubscriptionEdition = 4
    }

    // Content database generations from the build recorded in dbo.Versions.
    // Boundaries are public builds: SharePoint Server 2019 RTM 16.0.10337.12109
    // and the Subscription Edition preview 16.0.14131.20292; major 16 builds
    // before 2019 are SharePoint Server 2016.
    public static class SharePointBuilds
    {
        private static readonly Version Server2019 = new Version(16, 0, 10337, 12109);
        private static readonly Version SubscriptionEdition = new Version(16, 0, 14131, 20292);

        public static SharePointGeneration Classify(string build)
        {
            Version version;
            if (String.IsNullOrWhiteSpace(build) || !Version.TryParse(build, out version)) return SharePointGeneration.Unknown;
            return Classify(version);
        }
        public static SharePointGeneration Classify(Version build)
        {
            if (build == null) return SharePointGeneration.Unknown;
            if (build.Major < 16) return SharePointGeneration.SharePoint2013OrEarlier;
            if (build.Major > 16) return SharePointGeneration.Unknown;
            if (build >= SubscriptionEdition) return SharePointGeneration.SubscriptionEdition;
            if (build >= Server2019) return SharePointGeneration.SharePoint2019;
            return SharePointGeneration.SharePoint2016;
        }
        public static bool IsSupported(SharePointGeneration generation)
        {
            return generation == SharePointGeneration.SharePoint2016 || generation == SharePointGeneration.SharePoint2019 ||
                generation == SharePointGeneration.SubscriptionEdition;
        }
        public static string DisplayName(SharePointGeneration generation)
        {
            switch (generation)
            {
                case SharePointGeneration.SharePoint2013OrEarlier: return "SharePoint 2013 or earlier";
                case SharePointGeneration.SharePoint2016: return "SharePoint Server 2016";
                case SharePointGeneration.SharePoint2019: return "SharePoint Server 2019";
                case SharePointGeneration.SubscriptionEdition: return "SharePoint Server Subscription Edition";
                default: return "an unrecognized SharePoint version";
            }
        }
    }
}
