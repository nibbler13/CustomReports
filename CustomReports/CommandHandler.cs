using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace CustomReports {
	class CommandHandler : ICommand {
		private Action<object> action;
		private Predicate<object> canExecute;
		private event EventHandler CanExecuteChangedInternal;

		public event EventHandler CanExecuteChanged {
			add {
				CommandManager.RequerySuggested += value;
				CanExecuteChangedInternal += value;
			}
			remove {
				CommandManager.RequerySuggested -= value;
				CanExecuteChangedInternal -= value;
			}
		}

		public bool CanExecute(object parameter) {
			return canExecute != null && canExecute(parameter);
		}

		public void Execute(object parameter) {
			action(parameter);
		}

		public CommandHandler(Action<object> action) : this(action, DefaultCanExecute) { }

		public CommandHandler(Action<object> action, Predicate<object> canExecute) {
			this.action = action ?? throw new ArgumentNullException("action");
			this.canExecute = canExecute ?? throw new ArgumentNullException("canExecute");
		}

		public void OnCanExecuteChanged() {
			EventHandler eventHandler = CanExecuteChangedInternal;
			if (eventHandler != null)
				eventHandler.Invoke(this, EventArgs.Empty);
		}

		public void Destroy() {
			canExecute = _ => false;
			action = _ => { return; };
		}


		private static bool DefaultCanExecute(object parameter) {
			return true;
		}
	}
}
