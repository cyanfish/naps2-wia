namespace NAPS2.Wia
{
    /// <summary>
    /// Property value constants.
    /// </summary>
    public static class WiaPropertyValue
    {
        // Values for DOCUMENT_HANDLING_SELECT
        public const int FEEDER = 1;
        public const int FLATBED = 2;
        public const int DUPLEX = 4;
        public const int FRONT_FIRST = 8;
        public const int BACK_FIRST = 16;
        public const int FRONT_ONLY = 32;
        public const int BACK_ONLY = 64;
        public const int ADVANCED_DUPLEX = 1024;

        // Values for DOCUMENT_HANDLING_STATUS
        public const int FEED_READY = 1;

        // Values for PAGES
        public const int ALL_PAGES = 0;
    }
}