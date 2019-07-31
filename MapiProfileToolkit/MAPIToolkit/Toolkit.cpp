#include "pch.h"
#include "Toolkit.h"

Toolkit::Toolkit()
{
	m_profileWorker = new ProfileWorker();
	m_serviceWorker = new ServiceWorker();
	m_providerWorker = new ProviderWorker();
}

Toolkit::~Toolkit()
{
}
