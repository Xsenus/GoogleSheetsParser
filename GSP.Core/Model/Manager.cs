namespace GSP.Core.Model
{
    /// <summary>
    /// Менеджер.
    /// </summary>
    public class Manager
    {
        /// <summary>
        /// Имя менеджера.
        /// </summary>
        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }
}