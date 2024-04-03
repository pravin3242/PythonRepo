

paragraph = "EHR the ehr Eevolving healthcare landscape demands a strategic decision between off-the-shelf and custom solutions. While off-the-shelf options offer affordability and quick implementation, the rise in custom adoption signifies a shift towards tailored functionalities. Understanding your organization’s specific needs is key, ensuring the chosen solution aligns with workflows, scalability, and budget. The decision ultimately hinges on balancing immediate requirements with long-term goals for seamless, patient-centric care."
search_words = input("Enter the search words: ").lower().split()

words = paragraph.lower().split()

word_freq = {}
for word in search_words:
    word_frequency = words.count(word)
    word_freq[word] = word_frequency

for word, frequency in word_freq.items():
    print("The count of","'", word,"'","is",frequency)

