from django.shortcuts import get_object_or_404, redirect
from django.contrib import messages
from django.urls import reverse_lazy, reverse
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.views.decorators.http import require_POST
from django.db.models import Q
from ..models import RemoteMiningPlatform, Miner, Payout, Expense, TopUp
from ..forms import (
    RemoteMiningPlatformForm, MinerForm, PayoutForm, ExpenseForm, TopUpForm,
)


class PlatformListView(ListView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_list.html'
    context_object_name = 'platforms'
    paginate_by = 10


class PlatformDetailView(DetailView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_detail.html'
    context_object_name = 'platform'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_platform = self.get_object()
        
        # Navigate by name (matching list view order: alphabetical)
        # Previous = earlier in alphabet
        previous_platform = RemoteMiningPlatform.objects.filter(
            Q(name__lt=current_platform.name) |
            Q(name=current_platform.name, id__lt=current_platform.id)
        ).order_by('-name', '-id').first()
        
        # Next = later in alphabet
        next_platform = RemoteMiningPlatform.objects.filter(
            Q(name__gt=current_platform.name) |
            Q(name=current_platform.name, id__gt=current_platform.id)
        ).order_by('name', 'id').first()
        
        context['previous_platform'] = previous_platform
        context['next_platform'] = next_platform
        return context


class PlatformCreateView(CreateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    
    def get_success_url(self):
        return reverse_lazy('platform_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Platform created successfully.')
        return super().form_valid(form)


class PlatformUpdateView(UpdateView):
    model = RemoteMiningPlatform
    form_class = RemoteMiningPlatformForm
    template_name = 'mining/platform_form.html'
    
    def get_success_url(self):
        return reverse_lazy('platform_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Platform updated successfully.')
        return super().form_valid(form)


class PlatformDeleteView(DeleteView):
    model = RemoteMiningPlatform
    template_name = 'mining/platform_confirm_delete.html'
    success_url = reverse_lazy('platform_list')
    
    def delete(self, request, *args, **kwargs):
        messages.success(request, "Platform deleted successfully!")
        return super().delete(request, *args, **kwargs)


# Miner Views


class MinerListView(ListView):
    model = Miner
    template_name = 'mining/miner_list.html'
    context_object_name = 'miners'
    paginate_by = 50
    queryset = Miner.objects.select_related('platform')


class MinerDetailView(DetailView):
    model = Miner
    template_name = 'mining/miner_detail.html'
    context_object_name = 'miner'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_miner = self.get_object()
        
        # Get previous miner (lower ID)
        previous_miner = Miner.objects.filter(
            id__lt=current_miner.id
        ).order_by('-id').first()
        
        # Get next miner (higher ID)
        next_miner = Miner.objects.filter(
            id__gt=current_miner.id
        ).order_by('id').first()
        
        context['previous_miner'] = previous_miner
        context['next_miner'] = next_miner
        return context


class MinerCreateView(CreateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    
    def get_success_url(self):
        return reverse_lazy('miner_detail', kwargs={'pk': self.object.pk})

    def form_valid(self, form):
        messages.success(self.request, "Miner created successfully!")
        return super().form_valid(form)


class MinerUpdateView(UpdateView):
    model = Miner
    form_class = MinerForm
    template_name = 'mining/miner_form.html'
    
    def get_success_url(self):
        return reverse_lazy('miner_detail', kwargs={'pk': self.object.pk})

    def form_valid(self, form):
        messages.success(self.request, "Miner updated successfully!")
        return super().form_valid(form)


class MinerDeleteView(DeleteView):
    model = Miner
    template_name = 'mining/miner_confirm_delete.html'
    success_url = reverse_lazy('miner_list')
    context_object_name = 'miner'

    def delete(self, request, *args, **kwargs):
        messages.success(request, "Miner deleted successfully!")
        return super().delete(request, *args, **kwargs)




@require_POST
def toggle_miner_active(request, pk):
    """Toggle a miner's is_active status (on/off)"""
    miner = get_object_or_404(Miner, pk=pk)
    miner.is_active = not miner.is_active
    miner.save(update_fields=['is_active'])
    status = "ON" if miner.is_active else "OFF"
    messages.success(request, f"{miner.model} turned {status}.")
    return redirect('miner_detail', pk=pk)


# Payout Views


class PayoutListView(ListView):
    model = Payout
    template_name = 'mining/payout_list.html'
    context_object_name = 'payouts'
    paginate_by = 50
    queryset = Payout.objects.select_related('platform')


class PayoutDetailView(DetailView):
    model = Payout
    template_name = 'mining/payout_detail.html'
    context_object_name = 'payout'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_payout = self.get_object()
        
        # Navigate by payout_date (matching list view order: newest first)
        # Previous = newer payout (appears before in list)
        previous_payout = Payout.objects.filter(
            Q(payout_date__gt=current_payout.payout_date) |
            Q(payout_date=current_payout.payout_date, id__gt=current_payout.id)
        ).order_by('payout_date', 'id').first()
        
        # Next = older payout (appears after in list)
        next_payout = Payout.objects.filter(
            Q(payout_date__lt=current_payout.payout_date) |
            Q(payout_date=current_payout.payout_date, id__lt=current_payout.id)
        ).order_by('-payout_date', '-id').first()
        
        context['previous_payout'] = previous_payout
        context['next_payout'] = next_payout
        return context


class PayoutCreateView(CreateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    
    def get_success_url(self):
        return reverse_lazy('payout_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Payout added successfully!')
        return super().form_valid(form)


class PayoutUpdateView(UpdateView):
    model = Payout
    form_class = PayoutForm
    template_name = 'mining/payout_form.html'
    
    def get_success_url(self):
        return reverse_lazy('payout_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Payout updated successfully!')
        return super().form_valid(form)


class PayoutDeleteView(DeleteView):
    model = Payout
    template_name = 'mining/payout_confirm_delete.html'
    success_url = reverse_lazy('payout_list')
    context_object_name = 'payout'
    
    def delete(self, request, *args, **kwargs):
        messages.success(self.request, 'Payout deleted successfully!')
        return super().delete(request, *args, **kwargs)




class ExpenseListView(ListView):
    model = Expense
    template_name = 'mining/expense_list.html'
    context_object_name = 'expenses'
    paginate_by = 50
    queryset = Expense.objects.select_related('platform')


class ExpenseDetailView(DetailView):
    model = Expense
    template_name = 'mining/expense_detail.html'
    context_object_name = 'expense'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        current_expense = self.get_object()
        
        # Navigate by expense_date (matching list view order: newest first)
        # Previous = newer expense (appears before in list)
        previous_expense = Expense.objects.filter(
            Q(expense_date__gt=current_expense.expense_date) |
            Q(expense_date=current_expense.expense_date, id__gt=current_expense.id)
        ).order_by('expense_date', 'id').first()
        
        # Next = older expense (appears after in list)
        next_expense = Expense.objects.filter(
            Q(expense_date__lt=current_expense.expense_date) |
            Q(expense_date=current_expense.expense_date, id__lt=current_expense.id)
        ).order_by('-expense_date', '-id').first()
        
        context['previous_expense'] = previous_expense
        context['next_expense'] = next_expense
        return context


class ExpenseCreateView(CreateView):
    model = Expense
    form_class = ExpenseForm
    template_name = 'mining/expense_form.html'
    
    def get_success_url(self):
        return reverse_lazy('expense_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Expense created successfully!')
        return super().form_valid(form)


class ExpenseUpdateView(UpdateView):
    model = Expense
    form_class = ExpenseForm
    template_name = 'mining/expense_form.html'
    
    def get_success_url(self):
        return reverse_lazy('expense_detail', kwargs={'pk': self.object.pk})
    
    def form_valid(self, form):
        messages.success(self.request, 'Expense updated successfully!')
        return super().form_valid(form)


class ExpenseDeleteView(DeleteView):
    model = Expense
    template_name = 'mining/expense_confirm_delete.html'
    success_url = reverse_lazy('expense_list')
    context_object_name = 'expense'
    
    def delete(self, request, *args, **kwargs):
        messages.success(self.request, 'Expense deleted successfully!')
        return super().delete(request, *args, **kwargs)


# Dashboard Views


class TopUpListView(ListView):
    model = TopUp
    template_name = 'mining/topup_list.html'
    context_object_name = 'topups'
    paginate_by = 50
    queryset = TopUp.objects.select_related('platform')


class TopUpDetailView(DetailView):
    model = TopUp
    template_name = 'mining/topup_detail.html'
    context_object_name = 'topup'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        
        # Navigate by topup_date (matching list view order: newest first)
        topup = self.get_object()
        # Previous = newer top-up (appears before in list)
        context['previous_topup'] = TopUp.objects.filter(
            Q(topup_date__gt=topup.topup_date) |
            Q(topup_date=topup.topup_date, id__gt=topup.id)
        ).order_by('topup_date', 'id').first()
        
        # Next = older top-up (appears after in list)
        context['next_topup'] = TopUp.objects.filter(
            Q(topup_date__lt=topup.topup_date) |
            Q(topup_date=topup.topup_date, id__lt=topup.id)
        ).order_by('-topup_date', '-id').first()
        
        return context


class TopUpCreateView(CreateView):
    model = TopUp
    form_class = TopUpForm
    template_name = 'mining/topup_form.html'
    
    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(self.request, 'Top-Up created successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_detail', kwargs={'pk': self.object.pk})


class TopUpUpdateView(UpdateView):
    model = TopUp
    form_class = TopUpForm
    template_name = 'mining/topup_form.html'
    
    def form_valid(self, form):
        response = super().form_valid(form)
        messages.success(self.request, 'Top-Up updated successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_detail', kwargs={'pk': self.object.pk})


class TopUpDeleteView(DeleteView):
    model = TopUp
    template_name = 'mining/topup_confirm_delete.html'
    
    def delete(self, request, *args, **kwargs):
        response = super().delete(request, *args, **kwargs)
        messages.success(request, 'Top-Up deleted successfully!')
        return response
    
    def get_success_url(self):
        return reverse('topup_list')


