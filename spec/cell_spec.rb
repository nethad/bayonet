require 'spec_helper'

describe Bayonet::Cell do

  describe '#label' do

    it 'should concat row and column' do
      expect(Bayonet::Cell.new('AA', 999).label).to eq('AA999')
    end

    it 'should raise for invalid row and columns' do
      expect{ Bayonet::Cell.new('A1', 999).label}.to raise_error
      expect{ Bayonet::Cell.new('A', '1').label}.to raise_error
    end

  end

  describe '#valid?' do

    it 'valid rows' do
      expect(Bayonet::Cell.new('A', 1)).to be_valid
      expect(Bayonet::Cell.new('AA', 1)).to be_valid
      expect(Bayonet::Cell.new('Z', 1)).to be_valid
      expect(Bayonet::Cell.new('ZZ', 1)).to be_valid
    end

    it 'invalid rows' do
      expect(Bayonet::Cell.new('1', 1)).not_to be_valid
      expect(Bayonet::Cell.new('A1', 1)).not_to be_valid
      expect(Bayonet::Cell.new('1A', 1)).not_to be_valid
    end

    it 'valid columns' do
      expect(Bayonet::Cell.new('A', 1)).to be_valid
      expect(Bayonet::Cell.new('A', 999)).to be_valid
    end

    it 'invalid columns' do
      expect(Bayonet::Cell.new('A', -1)).not_to be_valid
      expect(Bayonet::Cell.new('A', 0)).not_to be_valid
      expect(Bayonet::Cell.new('A', 'A')).not_to be_valid
      expect(Bayonet::Cell.new('A', '1')).not_to be_valid
      expect(Bayonet::Cell.new('A', '1A')).not_to be_valid
    end

  end

end
