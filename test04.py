from django import forms
from django.shortcuts import render, redirect


# Создание формы
class ContactForm(forms.Form):
    name = forms.CharField(label='Ваше имя', max_length=100)
    email = forms.EmailField(label='Ваш email')
    message = forms.CharField(label='Сообщение', widget=forms.Textarea)


# Обработка формы в представлении (view)
def contact_view(request):
    if request.method == 'POST':
        form = ContactForm(request.POST)
        if form.is_valid():
            # Обработка данных
            name = form.cleaned_data['name']
            email = form.cleaned_data['email']
            message = form.cleaned_data['message']
            print(f"Имя: {name}, Email: {email}, Сообщение: {message}")
            return redirect('success_url')  # Перенаправление после успешной отправки
    else:
        form = ContactForm()

    return render(request, 'contact.html', {'form': form})