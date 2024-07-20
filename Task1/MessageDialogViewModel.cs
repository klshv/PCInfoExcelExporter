namespace Task1 {
    public sealed class MessageDialogViewModel {

        public MessageDialogViewModel(string message) {
            Message = message;
        }

        public string Message { get; }
    }
}