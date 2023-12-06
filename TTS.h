#pragma once
#include <regex>
#include <msclr/marshal_cppstd.h>
#include <iostream>
#include <sstream>
#include <windows.h>
#include <mmsystem.h>
#pragma comment(lib, "winmm.lib")

namespace TTS {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace System::Text::RegularExpressions;

	/// <summary>
	/// Summary for TTS
	/// </summary>
	public ref class TTS : public System::Windows::Forms::Form
	{
	public:
		TTS(void)
		{
			InitializeComponent();
			//
			//TODO: Add the constructor code here
			//
		}

	protected:
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~TTS()
		{
			if (components)
			{
				delete components;
			}
		}

	private: System::Windows::Forms::Button^ BrowsDoc;


	private: System::Windows::Forms::OpenFileDialog^ openFileDialog1;
	private: System::Windows::Forms::TextBox^ textBox1;
	private: System::Windows::Forms::Button^ speakButton;
	private: System::Windows::Forms::Button^ pauseButton;
	private: System::Windows::Forms::Button^ stopButton;




	// Create a SpeechSynthesizer instance
	System::Speech::Synthesis::SpeechSynthesizer speechSynthesizer;

	bool isPaused = false;
	private: System::ComponentModel::BackgroundWorker^ backgroundWorker1;
	private: System::Windows::Forms::TextBox^ textBox2;
	private: System::Windows::Forms::Button^ findButton;
			 int currentSearchIndex = -1;
	private: System::Windows::Forms::Button^ summarizeButton;
	private: System::Windows::Forms::PictureBox^ pictureBox1;


	protected:

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		void InitializeComponent(void)
		{
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(TTS::typeid));
			this->BrowsDoc = (gcnew System::Windows::Forms::Button());
			this->openFileDialog1 = (gcnew System::Windows::Forms::OpenFileDialog());
			this->textBox1 = (gcnew System::Windows::Forms::TextBox());
			this->speakButton = (gcnew System::Windows::Forms::Button());
			this->pauseButton = (gcnew System::Windows::Forms::Button());
			this->stopButton = (gcnew System::Windows::Forms::Button());
			this->backgroundWorker1 = (gcnew System::ComponentModel::BackgroundWorker());
			this->textBox2 = (gcnew System::Windows::Forms::TextBox());
			this->findButton = (gcnew System::Windows::Forms::Button());
			this->summarizeButton = (gcnew System::Windows::Forms::Button());
			this->pictureBox1 = (gcnew System::Windows::Forms::PictureBox());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->BeginInit();
			this->SuspendLayout();
			// 
			// BrowsDoc
			// 
			this->BrowsDoc->AccessibleDescription = L"use to browse your documents";
			this->BrowsDoc->AccessibleName = L"Browse Documents";
			this->BrowsDoc->Location = System::Drawing::Point(343, 80);
			this->BrowsDoc->Name = L"BrowsDoc";
			this->BrowsDoc->Size = System::Drawing::Size(115, 37);
			this->BrowsDoc->TabIndex = 1;
			this->BrowsDoc->Text = L"Browse Documents (Ctrl + O)";
			this->BrowsDoc->UseVisualStyleBackColor = true;
			this->BrowsDoc->Click += gcnew System::EventHandler(this, &TTS::button1_Click);
			// 
			// openFileDialog1
			// 
			this->openFileDialog1->FileName = L"openFileDialog1";
			// 
			// textBox1
			// 
			this->textBox1->Location = System::Drawing::Point(4, 89);
			this->textBox1->Name = L"textBox1";
			this->textBox1->Size = System::Drawing::Size(320, 20);
			this->textBox1->TabIndex = 2;
			this->textBox1->TextChanged += gcnew System::EventHandler(this, &TTS::textBox1_TextChanged);
			// 
			// speakButton
			// 
			this->speakButton->Location = System::Drawing::Point(4, 137);
			this->speakButton->Name = L"speakButton";
			this->speakButton->Size = System::Drawing::Size(115, 40);
			this->speakButton->TabIndex = 3;
			this->speakButton->Text = L"Speak (Ctrl + T)";
			this->speakButton->UseVisualStyleBackColor = true;
			this->speakButton->Click += gcnew System::EventHandler(this, &TTS::speakButton_Click);
			// 
			// pauseButton
			// 
			this->pauseButton->Location = System::Drawing::Point(176, 137);
			this->pauseButton->Name = L"pauseButton";
			this->pauseButton->Size = System::Drawing::Size(115, 40);
			this->pauseButton->TabIndex = 4;
			this->pauseButton->Text = L"Pause (Ctrl + P)";
			this->pauseButton->UseVisualStyleBackColor = true;
			this->pauseButton->Click += gcnew System::EventHandler(this, &TTS::pauseButton_Click);
			// 
			// stopButton
			// 
			this->stopButton->Location = System::Drawing::Point(343, 137);
			this->stopButton->Name = L"stopButton";
			this->stopButton->Size = System::Drawing::Size(115, 40);
			this->stopButton->TabIndex = 5;
			this->stopButton->Text = L"Stop (Ctrl + S)";
			this->stopButton->UseVisualStyleBackColor = true;
			this->stopButton->Click += gcnew System::EventHandler(this, &TTS::stopButton_Click);
			// 
			// textBox2
			// 
			this->textBox2->Location = System::Drawing::Point(4, 206);
			this->textBox2->Name = L"textBox2";
			this->textBox2->Size = System::Drawing::Size(320, 20);
			this->textBox2->TabIndex = 6;
			// 
			// findButton
			// 
			this->findButton->Location = System::Drawing::Point(343, 196);
			this->findButton->Name = L"findButton";
			this->findButton->Size = System::Drawing::Size(115, 38);
			this->findButton->TabIndex = 7;
			this->findButton->Text = L"Find (Ctrl + F)";
			this->findButton->UseVisualStyleBackColor = true;
			this->findButton->Click += gcnew System::EventHandler(this, &TTS::findButton_Click);
			this->findButton->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &TTS::TTS_KeyDown);
			// 
			// summarizeButton
			// 
			this->summarizeButton->Location = System::Drawing::Point(156, 239);
			this->summarizeButton->Name = L"summarizeButton";
			this->summarizeButton->Size = System::Drawing::Size(135, 38);
			this->summarizeButton->TabIndex = 8;
			this->summarizeButton->Text = L"Summarize Document (Ctrl + M)";
			this->summarizeButton->UseVisualStyleBackColor = true;
			this->summarizeButton->Click += gcnew System::EventHandler(this, &TTS::summarizeButton_Click);
			this->summarizeButton->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &TTS::TTS_KeyDown);
			// 
			// pictureBox1
			// 
			this->pictureBox1->Image = (cli::safe_cast<System::Drawing::Image^>(resources->GetObject(L"pictureBox1.Image")));
			this->pictureBox1->Location = System::Drawing::Point(4, 0);
			this->pictureBox1->Name = L"pictureBox1";
			this->pictureBox1->Size = System::Drawing::Size(216, 64);
			this->pictureBox1->SizeMode = System::Windows::Forms::PictureBoxSizeMode::StretchImage;
			this->pictureBox1->TabIndex = 9;
			this->pictureBox1->TabStop = false;
			// 
			// TTS
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(120)), static_cast<System::Int32>(static_cast<System::Byte>(232)),
				static_cast<System::Int32>(static_cast<System::Byte>(223)));
			this->ClientSize = System::Drawing::Size(470, 289);
			this->Controls->Add(this->pictureBox1);
			this->Controls->Add(this->summarizeButton);
			this->Controls->Add(this->findButton);
			this->Controls->Add(this->textBox2);
			this->Controls->Add(this->stopButton);
			this->Controls->Add(this->pauseButton);
			this->Controls->Add(this->speakButton);
			this->Controls->Add(this->textBox1);
			this->Controls->Add(this->BrowsDoc);
			this->KeyPreview = true;
			this->Name = L"TTS";
			this->Text = L"Libre Reader";
			this->KeyDown += gcnew System::Windows::Forms::KeyEventHandler(this, &TTS::TTS_KeyDown);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->pictureBox1))->EndInit();
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void label1_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		openFileDialog1->Filter = "Word,Text|*.doc;*.docx;*.txt";
		openFileDialog1->ShowDialog();
		textBox1->Text = openFileDialog1->FileName;
	}
	private: System::Void speakButton_Click(System::Object^ sender, System::EventArgs^ e) {
		try
		{
			// Read the content of the text file
			String^ filePath = textBox1->Text;
			String^ textContent = System::IO::File::ReadAllText(filePath);

			if (isPaused) {
				speechSynthesizer.Resume();
				isPaused = false;

				pauseButton->Text = "Pause";
			}

			speechSynthesizer.SpeakAsyncCancelAll();
			// Speak the content of the text file
			speechSynthesizer.SpeakAsync(textContent);
		}
		catch (Exception^ ex)
		{
			// Handle any exceptions (file not found, etc.)
			MessageBox::Show("An error occurred: " + ex->Message, "Error", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void pauseButton_Click(System::Object^ sender, System::EventArgs^ e) {
		if (isPaused) {
			speechSynthesizer.Resume();
			isPaused = false;

			pauseButton->Text = "Pause";
		}
		else {
			speechSynthesizer.Pause();
			isPaused = true;

			pauseButton->Text = "Resume";
		}
	}

	private: System::Void stopButton_Click(System::Object^ sender, System::EventArgs^ e) {
		speechSynthesizer.SpeakAsyncCancelAll();
	}


	private: System::Void textBox1_TextChanged(System::Object^ sender, System::EventArgs^ e) {
	}

	private: System::Void findButton_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ searchText = textBox2->Text;

		if (!String::IsNullOrEmpty(searchText)) {
			String^ filePath = textBox1->Text;

			if (!String::IsNullOrEmpty(filePath)) {
				try {
					String^ textContent = System::IO::File::ReadAllText(filePath);

					// Create a regex pattern for whole word matching (case-insensitive)
					String^ pattern = "\\b" + Regex::Escape(searchText) + "\\b";
					Regex^ regex = gcnew Regex(pattern, RegexOptions::IgnoreCase);

					// Find the starting match of the next occurrence of the search text
					Match^ match = regex->Match(textContent, currentSearchIndex + 1);

					if (match->Success) {
						// Update the current position
						currentSearchIndex = match->Index;

						// Extract the text from the found position to the end
						String^ newTextContent = textContent->Substring(match->Index);

						// Cancel any ongoing speech and speak the extracted text
						speechSynthesizer.SpeakAsyncCancelAll();
						speechSynthesizer.SpeakAsync(newTextContent);
					}
					else {
						// Reset the current position if no more occurrences are found
						currentSearchIndex = -1;

						// Play a message indicating no more instances found
						speechSynthesizer.SpeakAsyncCancelAll();
						playMP3("./sounds/system-error-notice-quite.mp3");
						System::Threading::Thread::Sleep(1500);
						String^ noMoreInstancesMessage = "No more occurance found.";
						speechSynthesizer.SpeakAsyncCancelAll();  // Cancel any ongoing speech
						speechSynthesizer.SpeakAsync(noMoreInstancesMessage);
					}
				}
				catch (Exception^ ex) {
					// Play a message for other types of errors
					playMP3("./sounds/system-error-notice-quite.mp3");
					System::Threading::Thread::Sleep(1500);
					String^ errorMessage = "An error occurred: " + ex->Message;
					speechSynthesizer.SpeakAsync(errorMessage);
				}
			}
			else {
				// Play a message indicating the file path is empty
				playMP3("./sounds/system-error-notice-quite.mp3");
				System::Threading::Thread::Sleep(1500);
				String^ emptyFilePathMessage = "File path is empty, please choose a file";
				speechSynthesizer.SpeakAsync(emptyFilePathMessage);
			}
		}
		else {
			// Play a message indicating no search text entered
			playMP3("./sounds/system-error-notice-quite.mp3");
			System::Threading::Thread::Sleep(1500);
			String^ noSearchTextMessage = "Please enter text to find.";
			speechSynthesizer.SpeakAsync(noSearchTextMessage);
		}
	}

	void playMP3(System::String^ mp3Path) {
		mciSendString(L"close mp3", NULL, 0, NULL);
		System::String^ command = "open \"" + mp3Path + "\" type mpegvideo alias mp3";
		mciSendString((const wchar_t*)System::Runtime::InteropServices::Marshal::StringToHGlobalUni(command).ToPointer(), NULL, 0, NULL);
		mciSendString(L"play mp3", NULL, 0, NULL);
	}

	private: System::Void summarizeButton_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ scriptPath = "./summarize_script.py";
		String^ documentPath = textBox1->Text;

		// Speak "Summarizing"
		speechSynthesizer.SpeakAsyncCancelAll();
		playMP3("./sounds/announcement-sound.mp3");
		System::Threading::Thread::Sleep(3000);
		speechSynthesizer.SpeakAsync("Summarizing your document, please wait");

		System::Diagnostics::Process^ process = gcnew System::Diagnostics::Process();
		process->StartInfo->FileName = "C:\\Users\\nguye\\AppData\\Local\\Programs\\Python\\Python38\\python.exe";  // Change this to the path to your Python 3.8 interpreter
		process->StartInfo->Arguments = String::Format("\"{0}\" \"{1}\"", scriptPath, documentPath);
		process->StartInfo->RedirectStandardOutput = true;
		process->StartInfo->UseShellExecute = false;
		process->StartInfo->CreateNoWindow = true;

		process->Start();

		// Read the output of the Python script
		String^ summaryText = process->StandardOutput->ReadToEnd();

		process->WaitForExit();

		// Speak the summary
		speechSynthesizer.SpeakAsyncCancelAll();
		playMP3("./sounds/announcement-sound.mp3");
		System::Threading::Thread::Sleep(3000);
		speechSynthesizer.SpeakAsync("Summarizing completed, here is your summarize");
		speechSynthesizer.SpeakAsync(summaryText);

		// Save the summary to a text file
		try {
			String^ summaryFilePath = "./summary.txt";
			System::IO::File::WriteAllText(summaryFilePath, summaryText);
			// Comment out the following line to stop the pop-up message
			// MessageBox::Show("Summary saved to: " + summaryFilePath, "Summary Saved", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		catch (Exception^ ex) {
			String^ errorMessage = "An error occurred while saving the summary: " + ex->Message;
			speechSynthesizer.SpeakAsync(errorMessage);
		}
	}

	private: System::Void TTS_KeyDown(System::Object^ sender, System::Windows::Forms::KeyEventArgs^ e) {
		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::O)
		{
			openFileDialog1->Filter = "Word,Text|*.doc;*.docx;*.txt";
			openFileDialog1->ShowDialog();
			textBox1->Text = openFileDialog1->FileName;
		}

		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::T) {
			speakButton_Click(sender, e);
		}

		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::P) {
			pauseButton_Click(sender, e);
		}

		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::S) {
			stopButton_Click(sender, e);
		}

		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::F) {
			findButton_Click(sender, e);
		}

		if (e->Control && e->KeyCode == System::Windows::Forms::Keys::M) {
			summarizeButton_Click(sender, e);
		}
	}

};
}
