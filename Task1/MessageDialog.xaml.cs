namespace Task1 {
    
    public partial class MessageDialog {
        public MessageDialog() {
            InitializeComponent();
            
            var viewModel = new MessageDialogViewModel("Message");
            DataContext = viewModel;
        }
    }
}